using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Cryptography;
using System.Text;

namespace FeiNuo.Core.Security
{
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
            var credentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256);

            var jwtToken = new JwtSecurityToken(
                issuer: cfg.Jwt.Issuer,
                audience: cfg.Jwt.Audience,
                signingCredentials: credentials,
                claims: user.UserClaims,
                expires: DateTime.Now.AddSeconds(cfg.TokenExpiration)
            );

            var token = tokenHandler.WriteToken(jwtToken);
            return Task.FromResult(token);
        }

        /// <summary>
        /// 验证Token合法性，通过后根据token获取用户信息
        /// </summary>
        public async Task<TokenValidationResult> ValidateTokenAsync(string token)
        {
            var validationParameters = new TokenValidationParameters()
            {
                // 验证发行人
                ValidateIssuer = !string.IsNullOrEmpty(cfg.Jwt.Issuer),
                ValidIssuer = cfg.Jwt.Issuer,

                // 验证受众人
                ValidateAudience = !string.IsNullOrEmpty(cfg.Jwt.Audience),
                ValidAudience = cfg.Jwt.Audience,

                // 验证签名
                ValidateIssuerSigningKey = true,
                IssuerSigningKey = new SymmetricSecurityKey(SHA256.HashData(Encoding.UTF8.GetBytes(cfg.Jwt.SigningKey))),

                RequireExpirationTime = cfg.TokenExpiration > 0,

                // 允许服务器时间偏移量(默认1800秒)
                // 即我们配置的过期时间加上这个允许偏移的时间值，才是真正过期的时间(过期时间 +偏移值)
                ClockSkew = TimeSpan.FromSeconds(cfg.Jwt.ClockSkew),
            };

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
                var refresh = DateTime.Now > result.SecurityToken.ValidTo.ToLocalTime();

                return new TokenValidationResult(user) { RefreshToken = refresh };
            }
            else return new TokenValidationResult(result.Exception.Message);
        }

        public async Task DisableTokenAsync(string token)
        {
            // jwt 无法作废，通过加入黑名单的方式实现
            var jwtToken = tokenHandler.ReadJwtToken(token);
            // HACK 这里依赖分布式缓存，如果没有的话，默认内存缓存，重启后会丢失
            await cache.SetStringAsync(GetDisableTokenKey(token), "1", new DistributedCacheEntryOptions()
            {
                AbsoluteExpiration = jwtToken.ValidTo.ToLocalTime(),
            });
        }

        private static string GetDisableTokenKey(string token)
        {
            return AppConstants.CACHE_PREFIX_FORBIDDEN + token;
        }
    }
}