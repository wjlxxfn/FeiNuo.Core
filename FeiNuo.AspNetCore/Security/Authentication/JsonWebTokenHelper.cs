using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Tokens;
using System.Security.Cryptography;
using System.Text;

namespace FeiNuo.AspNetCore.Security.Authentication;

public class JsonWebTokenHelper
{
    public static TokenValidationParameters GetTokenValidationParameters(SecurityOptions cfg)
    {
        return new TokenValidationParameters()
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

            // 允许服务器时间偏移量(默认300秒)
            // 即我们配置的过期时间加上这个允许偏移的时间值，才是真正过期的时间(过期时间 +偏移值)
            ClockSkew = TimeSpan.FromSeconds(cfg.Jwt.ClockSkew),

            NameClaimType = FNClaimTypes.UserName,
            RoleClaimType = FNClaimTypes.Role
        };
    }

    public static JwtBearerEvents GetJwtBearerEvents(SecurityOptions cfg)
    {
        return new JwtBearerEvents()
        {
            // 没有登录或token过期等 401
            OnChallenge = async context =>
            {
                context.HandleResponse();
                var error = "请先登录系统!" + (context.Error ?? "");
                var detail = context.ErrorDescription;

                var objMsg = new MessageResult(error, MessageTypeEnum.Warning, detail);
                context.Response.StatusCode = StatusCodes.Status401Unauthorized;
                context.Response.ContentType = "application/json; charset=utf-8";
                await context.Response.WriteAsJsonAsync(objMsg);
            },
            // 没权限 403
            OnForbidden = async context =>
            {
                var resp = new MessageResult("没有操作权限", MessageTypeEnum.Warning, context.Request.GetDisplayUrl());
                context.Response.StatusCode = StatusCodes.Status403Forbidden;
                context.Response.ContentType = "application/json; charset=utf-8";
                await context.Response.WriteAsJsonAsync(resp);
            },
            // token通过后验证是否作废，是否快到期等
            OnTokenValidated = async context =>
            {
                var token = (JsonWebToken)context.SecurityToken;

                // 判断是否已退出登录
                if (cfg.Jwt.CheckForbidden)
                {
                    var cache = context.HttpContext.RequestServices.GetRequiredService<IDistributedCache>();
                    var forbid = await cache.GetStringAsync(AppConstants.CACHE_PREFIX_FORBIDDEN + token.EncodedSignature);
                    if (forbid != null)
                    {
                        context.Fail("token已作废");
                        return;
                    }
                }

                // 到期了，但是在缓冲期内，生成新token,加入到响应头里，前端替换token
                var checkTime = cfg.TokenExpiration > (cfg.Jwt.ClockSkew * 2) ? (cfg.Jwt.ClockSkew * -1) : 0;
                if (DateTime.Now > token.ValidTo.ToLocalTime().AddSeconds(checkTime))
                {
                    var user = new LoginUser(token.Claims);
                    var tokenService = context.HttpContext.RequestServices.GetRequiredService<ITokenService>();
                    var newToken = await tokenService.CreateTokenAsync(user);
                    context.Response.Headers.Append(AppConstants.REFRESH_TOKEN_KEY, newToken);
                }
            },
        };
    }
}
