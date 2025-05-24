global using FeiNuo.Core;
using FeiNuo.AspNetCore.Mvc;
using FeiNuo.AspNetCore.Security;
using FeiNuo.AspNetCore.Security.Authentication;
using FeiNuo.AspNetCore.Security.Authorization;
using FeiNuo.Core.Login;
using FeiNuo.Core.Utilities;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.IdentityModel.JsonWebTokens;
using System.ComponentModel;

namespace FeiNuo.AspNetCore;

/// <summary>
/// 服务注入扩展
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// 依次注入：[Service]标注的服务，JWT认证,权限策略,控制器
    /// </summary>
    /// <param name="useJwtAuthentication">true，使用jwt认证；false:使用自定义的token认证，默认true</param>
    /// <remarks>
    /// useJwtAuthentication=false时，系统会根据注入的ITokenService类型判断是jwt还是cache的实现，默认注入的是CacheTokenService
    /// </remarks>
    public static IServiceCollection AddFNAspNetCore(this IServiceCollection services, IConfiguration configuration, bool useJwtAuthentication = true)
    {
        // 服务
        services.AddFNServices();

        // 认证
        if (useJwtAuthentication)
        {
            services.AddFNAuthenticationJwt(configuration);
        }
        else
        {
            services.AddFNAuthenticationToken(configuration);
        }

        // 授权
        services.AddFNAuthorization(configuration);

        // 控制器
        services.AddFNControllers();

        return services;
    }

    #region 注入服务类：BaseService的子类以及[ServiceAttribute]
    /// <summary>
    /// 注入各项服务：服务层的服务类，注解注入的；注入内存缓存服务；注入默认的日志记录服务
    /// </summary>
    public static IServiceCollection AddFNServices(this IServiceCollection services)
    {
        // 自动注入[Service]特性的类或继承自BaseService接口的类
        services.AutoInjectServcice();

        // 注入内存缓存和默认的分布式缓存
        services.AddMemoryCache();
        services.AddDistributedMemoryCache();



        return services;
    }
    #endregion

    #region 注入Controller：实现异常处理，自定义验证失败后的返回内容，JSON配置
    /// <summary>
    /// 注入控制器：services.AddControllers
    /// </summary>
    /// <remarks>
    /// 1、添加自定义异常处理过滤器 <br/>
    /// 2、自定义模型验证失败后的返回内容 <br/>
    /// 3、配置Json格式化规则
    /// </remarks>
    public static IServiceCollection AddFNControllers(this IServiceCollection services)
    {
        services.AddControllers(config =>
        {
            // 添加自定义异常处理过滤器
            config.Filters.Add<AppExceptionFilter>();
            //  DateOnly，默认只支持日期格式，如果前端是完整的日期包含时间的转换不了,这里添加一个转换类
            TypeDescriptor.AddAttributes(typeof(DateOnly), new TypeConverterAttribute(typeof(DateOnlyTypeConverter)));
        })
        .ConfigureApiBehaviorOptions(options =>
        {
            // 数据校验不通过返回400，修改下返回格式
            options.InvalidModelStateResponseFactory = (context) =>
            {
                if (context.HttpContext.Request.Method == HttpMethod.Patch.Method)
                {
                    return null!;
                }

                var errors = context.ModelState.Values
                .Select(a => string.Join(",", a.Errors.Select(t => t.ErrorMessage)))
                .Where(a => !string.IsNullOrEmpty(a))
                .ToArray();
                var msgVo = new MessageResult("数据效验不通过：" + string.Join("，", errors), MessageTypeEnum.Error);
                // 400 
                return new BadRequestObjectResult(msgVo);
            };
        })
        // 配置Json格式化规则
        .AddJsonOptions(options => JsonUtils.MergeSerializerOptions(options.JsonSerializerOptions));
        return services;
    }
    #endregion

    #region 认证，授权。认证分jwt和token两种方式
    /// <summary>
    /// Jwt 认证
    /// </summary>
    public static IServiceCollection AddFNAuthenticationJwt(this IServiceCollection services, IConfiguration configuration)
    {
        // 注入token服务
        services.TryAddSingleton<ITokenService, JwtTokenService>();

        // 添加初始的登录用户服务，保证新初始化项目时不报错。
        services.TryAddScoped<ILoginUserService, SimpleLoginUserService>();

        // 注入登录服务
        services.TryAddScoped<ILoginService, LoginService>();

        // 调整role对应的claimType,这里需要清空默认的映射，否则不会使用自定义的RoleClaimType
        JsonWebTokenHandler.DefaultInboundClaimTypeMap.Clear();
        var cfg = configuration.GetSection(SecurityOptions.ConfigKey).Get<SecurityOptions>() ?? new();
        services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme).AddJwtBearer(options =>
        {
            options.TokenValidationParameters = JsonWebTokenHelper.GetTokenValidationParameters(cfg);
            // jwt事件处理：401，403，以及验证通过后刷新token等
            options.Events = JsonWebTokenHelper.GetJwtBearerEvents(cfg);
        });

        return services;
    }

    /// <summary>
    /// Token 认证，默认注入CacheTokenService
    /// </summary>
    public static IServiceCollection AddFNAuthenticationToken(this IServiceCollection services, IConfiguration configuration)
    {
        // 注入token服务
        services.TryAddSingleton<ITokenService, CacheTokenService>();

        // 添加初始的登录用户服务，保证新初始化项目时不报错。
        services.TryAddScoped<ILoginUserService, SimpleLoginUserService>();
        // 注入登录服务
        services.TryAddScoped<ILoginService, LoginService>();

        var scheme = JwtBearerDefaults.AuthenticationScheme;
        services.AddAuthentication(scheme).AddScheme<AuthenticationSchemeOptions, TokenAuthenticationHandler>(scheme, null);

        return services;
    }

    /// <summary>
    /// 鉴权：超管策略，忽略策略。 Permission特性
    /// </summary>
    /// <param name="setFallbackPolicy">true:添加默认策略，必须登录，默认为true</param>
    public static IServiceCollection AddFNAuthorization(this IServiceCollection services, IConfiguration configuration, bool setFallbackPolicy = true)
    {
        // 授权
        var builder = services.AddAuthorizationBuilder()
            // 超管账号策略（只有SuperAdmin才行，角色加上SuperAdmin也不行）
            .AddPolicy(AppConstants.SUPER_ADMIN, c => c.RequireUserName(AppConstants.SUPER_ADMIN))
            // 不需要授权的策略
            .AddPolicy(AppConstants.AUTH_POLICY_IGNORE, c => c.RequireAssertion(v => true));

        // 默认策略,必须登录
        if (setFallbackPolicy)
        {
            builder.SetFallbackPolicy(new AuthorizationPolicyBuilder().RequireAuthenticatedUser().Build());
        }

        // 授权：添加超管策略，允许所有权限
        services.AddSingleton<IAuthorizationHandler, SuperAdminAuthorizationHandler>();

        // 授权：通过permission字符串授权[Permission("system:user:delete")]
        services.AddSingleton<IAuthorizationHandler, PermissionAuthorizationHandler>();

        return services;
    }
    #endregion
}
