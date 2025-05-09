using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;

namespace FeiNuo.AspNetCore.Security.Authentication;

/// <summary>
/// JWT Token服务类
/// </summary>
public class JwtTokenService : ITokenService
{
    private readonly IDistributedCache cache;
    private readonly SecurityOptions cfg = new();
    private readonly JwtSecurityTokenHandler tokenHandler = new();

    /// <summary>
    /// 构造函数：注入JWT配置参数
    /// </summary>
    public JwtTokenService(IConfiguration configuration, IDistributedCache cache)
    {
        this.cache = cache;
        configuration.GetSection(SecurityOptions.ConfigKey).Bind(cfg);
    }

    /// <summary>
    /// 根据用户信息创建Token
    /// </summary>
    public Task<string> CreateTokenAsync(LoginUser user)
    {
        var securityKey = new SymmetricSecurityKey(SHA256.HashData(Encoding.UTF8.GetBytes(cfg.Jwt.SigningKey)));
        var tokenDescriptor = new SecurityTokenDescriptor
        {
            Issuer = cfg.Jwt.Issuer,
            Audience = cfg.Jwt.Audience,
            Subject = new ClaimsIdentity(user.UserClaims),
            Expires = DateTime.UtcNow.AddSeconds(cfg.TokenExpiration),
            SigningCredentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256),
        };

        var token = tokenHandler.CreateToken(tokenDescriptor);
        return Task.FromResult(tokenHandler.WriteToken(token));
    }

    /// <summary>
    /// 验证Token合法性，通过后根据token获取用户信息
    /// </summary>
    public async Task<TokenValidationResult> ValidateTokenAsync(string token)
    {
        var validationParameters = JsonWebTokenHelper.GetTokenValidationParameters(cfg);
        var result = await tokenHandler.ValidateTokenAsync(token, validationParameters);

        if (result.IsValid)
        {
            // 作废的，退出的
            var zf = await cache.GetStringAsync(GetDisableTokenKey(token));
            if (zf != null)
            {
                return new TokenValidationResult("TOKEN已作废");
            }
            var user = new LoginUser(((JwtSecurityToken)result.SecurityToken).Claims);

            // 判断是否需要刷新, 虽然过期但在缓冲期内仍然可以访问，这时刷新下token
            var checkTime = cfg.TokenExpiration > (cfg.Jwt.ClockSkew * 2) ? (cfg.Jwt.ClockSkew * -1) : 0;
            var refresh = DateTime.Now > result.SecurityToken.ValidTo.ToLocalTime().AddSeconds(checkTime);
            return new TokenValidationResult(user) { RefreshToken = refresh };
        }
        else return new TokenValidationResult(result.Exception.Message);
    }

    public async Task DisableTokenAsync(string token)
    {
        // jwt 无法作废，通过加入黑名单的方式实现
        if (cfg.Jwt.CheckForbidden)
        {
            var jwtToken = tokenHandler.ReadJwtToken(token);
            // 这里依赖分布式缓存，如果没有的话，默认内存缓存，重启后会丢失
            await cache.SetStringAsync(GetDisableTokenKey(jwtToken.RawSignature), "1", new DistributedCacheEntryOptions()
            {
                AbsoluteExpiration = jwtToken.ValidTo.ToLocalTime(),
            });
        }
    }

    private static string GetDisableTokenKey(string token)
    {
        return AppConstants.CACHE_PREFIX_FORBIDDEN + token;
    }
}