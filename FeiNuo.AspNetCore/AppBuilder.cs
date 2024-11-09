global using FeiNuo.Core;
using FeiNuo.AspNetCore.Security;
using FeiNuo.AspNetCore.Security.Authentication;
using FeiNuo.AspNetCore.Security.Authorization;
using FeiNuo.AspNetCore.Security.FormLogin;
using FeiNuo.Core.Utilities;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.OpenApi.Models;
using System.ComponentModel;

namespace FeiNuo.AspNetCore;

/// <summary>
/// 服务注入扩展
/// </summary>
public static class ServiceCollectionExtensions
{
    #region 注入服务类：BaseService的子类以及[ServiceAttribute]
    /// <summary>
    /// 注入各项服务：服务层的服务类，注解注入的，以及其他需要注入的组件
    /// </summary>
    public static IServiceCollection AddAppServices(this IServiceCollection services)
    {
        // 注入Http上下文类，服务类中有时需要使用
        services.AddHttpContextAccessor();

        // 自动注入[Service]特性的类或实现IService接口的类
        services.AutoInjectServcice();

        // 注入内存缓存和默认的分布式缓存
        services.AddMemoryCache();
        services.AddDistributedMemoryCache();

        // 添加默认的操作日志记录服务
        services.TryAddSingleton<ILogService, SimpleLogService>();

        return services;
    }
    #endregion

    #region 注入控制器：实现异常处理，自定义验证失败后的返回内容，JSON配置
    /// <summary>
    /// 注入控制器：异常处理，自定义验证失败后的返回内容，JSON配置
    /// </summary>
    public static IServiceCollection AddAppControllers(this IServiceCollection services)
    {
        services.AddControllers(config =>
        {
            // 添加自定义异常处理过滤器
            config.Filters.Add<AppExceptionFilter>();
            TypeDescriptor.AddAttributes(typeof(DateOnly), new TypeConverterAttribute(typeof(DateOnlyTypeConverter)));
        })
        .ConfigureApiBehaviorOptions(options =>
        {
            // 数据校验不通过返回400，修改下返回格式
            options.InvalidModelStateResponseFactory = (context) =>
            {
                var errors = context.ModelState.Values
                    .Select(a => string.Join(",", a.Errors.Select(t => t.ErrorMessage)))
                    .Where(a => !string.IsNullOrEmpty(a))
                    .ToArray();
                var msgVo = new MessageResult("数据效验不通过：" + string.Join("，", errors), MessageType.Error);
                // 400 
                return new BadRequestObjectResult(msgVo);
            };
        })
        // 配置Json格式化规则
        .AddJsonOptions(options => JsonUtils.MergeSerializerOptions(options.JsonSerializerOptions));
        return services;
    }
    #endregion

    #region 认证，授权
    /// <summary>
    /// 添加Token认证
    /// </summary>
    public static IServiceCollection AddAppSecurity(this IServiceCollection services, IConfiguration configuration)
    {
        #region 注入 TokenService 默认使用jwt，
        var cfg = configuration.GetSection(SecurityOptions.ConfigKey).Get<SecurityOptions>() ?? new();
        if (cfg.TokenType != "Jwt" && cfg.TokenType != "Cache")
        {
            throw new ArgumentException("Token类型只允许配置:Jwt、Cache");
        }
        if (cfg.TokenType == "Jwt")
        {
            services.TryAddSingleton<ITokenService, JwtTokenService>();
        }
        else if (cfg.TokenType == "Cache")
        {
            services.TryAddSingleton<ITokenService, CacheTokenService>();
        }
        #endregion

        // 添加初始的登录用户服务，保证新初始化项目时不报错。且在使用登录接口报错时给出提示需实现ILoginUserService
        services.TryAddScoped<ILoginUserService, SimpleLoginUserService>();
        // 注入登录服务
        services.TryAddScoped<ILoginService, LoginService>();

        // 认证
        var scheme = TokenAuthentication.AuthenticationScheme;
        services.AddAuthentication(scheme).AddScheme<AuthenticationSchemeOptions, TokenAuthentication>(scheme, null);

        // 授权
        services.AddAuthorizationBuilder()
            // 超管账号策略（只有SuperAdmin才行，角色加上SuperAdmin也不行）
            .AddPolicy(AppConstants.SUPER_ADMIN, c => c.RequireUserName(AppConstants.SUPER_ADMIN))
            // 默认策略,必须登录
            .SetFallbackPolicy(new AuthorizationPolicyBuilder().RequireAuthenticatedUser().Build());

        // 授权：添加超管策略，允许所有权限
        services.AddSingleton<IAuthorizationHandler, SuperAdminAuthorizationHandler>();
        // 授权：通过permission字符串授权[Permission("system:user:delete")]
        services.AddSingleton<IAuthorizationHandler, PermissionAuthorizationHandler>();

        #region 官方jwt认证扩展，原理类似，这里不再使用
        //services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme).AddJwtBearer(options =>
        //{
        //    options.TokenValidationParameters = new TokenValidationParameters()
        //    {
        //        // 验证发行人
        //        ValidateIssuer = !string.IsNullOrEmpty(cfg.Jwt.Issuer),
        //        ValidIssuer = cfg.Jwt.Issuer,
        //        // 验证受众人
        //        ValidateAudience = !string.IsNullOrEmpty(cfg.Jwt.Audience),
        //        ValidAudience = cfg.Jwt.Audience,

        //        // 验证签名
        //        ValidateIssuerSigningKey = true,
        //        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(cfg.Jwt.SigningKey)),

        //        RequireExpirationTime = cfg.TokenExpiration > 0,

        //        // 允许服务器时间偏移量(默认300秒)
        //        // 即我们配置的过期时间加上这个允许偏移的时间值，才是真正过期的时间(过期时间 +偏移值)
        //        ClockSkew = TimeSpan.FromSeconds(cfg.RefreshInterval),
        //    };

        //    // 可根据 JwtBearerHandler 源码查看默认的处理逻辑，这里给401，403添加统一的响应
        //    options.Events = new JwtBearerEvents()
        //    {
        //        // 没有登录或token过期等 401
        //        OnChallenge = async context =>
        //        {
        //            context.HandleResponse();

        //            var error = "请先登录" + (string.IsNullOrEmpty(context.Error) ? "" : (":" + context.Error));
        //            var detail = context.ErrorDescription;
        //            var objMsg = new ResponseMessage(error, MessageType.Warning, detail);

        //            context.Response.StatusCode = StatusCodes.Status401Unauthorized;
        //            context.Response.ContentType = "application/json; charset=utf-8";
        //            await context.Response.WriteAsJsonAsync(objMsg);
        //        },
        //        // 没权限 403
        //        OnForbidden = async context =>
        //        {
        //            var resp = new ResponseMessage("没有操作权限", MessageType.Warning);

        //            context.Response.StatusCode = StatusCodes.Status403Forbidden;
        //            context.Response.ContentType = "application/json; charset=utf-8";
        //            await context.Response.WriteAsJsonAsync(resp);
        //        },
        //        // token通过后验证是否作废，是否快到期等
        //        OnTokenValidated = async context =>
        //        {
        //            JwtSecurityToken token = (JwtSecurityToken)context.SecurityToken;

        //            var user = new LoginUser(token.Claims, true);

        //            // 达到刷新时间，生成新token,加入到响应头里，前端替换token
        //            if (cfg.RefreshInterval > 0 && ((DateTime.Now - token!.IssuedAt).TotalSeconds > cfg.RefreshInterval))
        //            {
        //                var tokenService = context.HttpContext.RequestServices.GetRequiredService<ITokenService>();
        //                var newToken = await tokenService.CreateTokenAsync(user);
        //                context.Response.Headers.Append("Authorization", newToken);
        //            }

        //            // 认证信息(用户名，角色)加入到Pricipal中,可以使用User.Identify.Name或User.IsInRole
        //            var identify = context.Principal!.Identity as ClaimsIdentity;
        //            identify?.AddClaims(user.Claims);
        //        },
        //    };
        //});
        #endregion

        return services;
    }
    #endregion

    #region Swagger
    public static IServiceCollection AddAppSwaggerGen(this IServiceCollection services, IConfiguration configuration, string[] xmlPaths)
    {
        // 注入 Swagger
        services.AddEndpointsApiExplorer();
        services.AddSwaggerGen(c =>
        {
            var scheme = TokenAuthentication.AuthenticationScheme;
            c.AddSecurityDefinition(scheme, new OpenApiSecurityScheme
            {
                Description = $"授权认证，在下方输入Bearer {{token}},例如：<br/>Bearer {AppConstants.SUPER_ADMIN_TOKEN}",
                Name = "Authorization",
                In = ParameterLocation.Header,
                Type = SecuritySchemeType.ApiKey,
            });
            c.AddSecurityRequirement(new OpenApiSecurityRequirement
            {
                {
                    new OpenApiSecurityScheme {Reference = new OpenApiReference { Type = ReferenceType.SecurityScheme, Id = scheme }},
                    Array.Empty<string>()
                }
            });
            var basePath = Path.GetDirectoryName(AppContext.BaseDirectory)!;
            foreach (var path in xmlPaths)
            {
                c.IncludeXmlComments(Path.Combine(basePath, path), true);
            }
        });
        return services;

    }
    #endregion
}
