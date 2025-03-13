# FeiNuo.AspNetCore
## 功能介绍
> 本项目基于Net8开发，用于辅助开发WebApi项目。    
1. 自动注入服务
2. 认证授权服务
3. MVC相关辅助功能
4. 其他辅助功能

## 一、自动注入服务
  1、_自动注入[Service]特性的类或继承自BaseService的类_
  ```
    builder.Services.AutoInjectServcice();

    对于标注有[Service]特性的类：
    1. 如果特性中有注入类型，则以特性中指定的类型为准
    2. 如果特性中没有注入类型，则根据当前类实现的接口类型注入
    3. 如果当前类也没有实现接口，则根据当前类的类型注入
    4. 生命周期默认为 `ServiceLifetime.Scoped`，可以在特性中修改

    实现BaseService的所有类
    1. 如果该类有[Service]特性,按前面规则注入
    2. 有实现的接口(BaseService除外)则根据接口类型注入
    3. 如果没有实现接口(BaseService除外)，则根据当前类型注入
    4. 注入的生命周期全部为 `ServiceLifetime.Scoped`
    5. 如果默认注入的不合适，可添加[Service]特性自定义
  ```

  2、_注入常用服务类,自动注入+缓存服务+日志服务_
  ```
    public static IServiceCollection AddFNServices(this IServiceCollection services)
    {
        // 自动注入[Service]特性的类或继承自BaseService接口的类
        services.AutoInjectServcice();

        // 注入内存缓存和默认的分布式缓存
        services.AddMemoryCache();
        services.AddDistributedMemoryCache();

        // 添加默认的操作日志记录服务
        services.TryAddSingleton<ILogService, SimpleLogService>();

        return services;
    }
   ```

  3、_注入控制器，实现异常处理，JSON配置_
   ```
    public static IServiceCollection AddFNControllers(this IServiceCollection services)
    {
        services.AddControllers(config =>
        {
            // 添加自定义异常处理过滤器
            // 异常响应：
            // 1. 没有登录信息：返回 401,UnauthorizedResult
            // 2. 没有权限：返回 403,ForbidResult
            // 3. 模型数据验证不通过：返回 400,BadRequestResult
            // 4. 系统自定义异常消息：返回 422,UnprocessableEntityResult
            // 5. 系统异常：返回 500
            config.Filters.Add<AppExceptionFilter>();
            //  DateOnly，默认只支持日期格式，如果前端是完整的日期包含时间的转换不了,这里添加一个转换类
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
                var msgVo = new MessageResult("数据效验不通过：" + string.Join("，", errors), MessageTypeEnum.Error);
                // 400 
                return new BadRequestObjectResult(msgVo);
            };
        })
        // 配置Json格式化规则
        // 1. 统一序列化配置
        // 2. null值默认不输出
        // 3. 阻断循环引用
        // 4. 属性转换策略：小驼峰
        // 5. 支持从字符串转成数字
        // 6. 日期类型默认输出格式 yyyy-MM-dd HH:mm:ss
        .AddJsonOptions(options => JsonUtils.MergeSerializerOptions(options.JsonSerializerOptions));
        return services;
    }
   ```

  4、注入用户认证，实现登录服务
   ```
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
   ```
  5、注入权限管理
  ```
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
  ```

  6、全部注入
   ```
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
 ```


    
## 三、认证授权相关实现

### 1、登录接口
    提供登录相关Rest接口: LoginController
    需要注入查询登录用户的实现类：ILoginUserService  

|    接口   | 类型  |     参数   |   返回  |
|    ----   | :---: |    ------  |   ----  |
| /login    | POST  | LoginForm  | Token字符串
| /logout   | POST  | 无         | 无
| /userinfo | GET   | 无         | LoginUser
| /captcha  | GET   | 无         | CaptchaResult  

#### 必需实现 ILoginUserService，以提供登录所需的用户信息
```
public interface ILoginUserService
{
    /// <summary>
    /// 通过用户名查询用户信息：包括，用户名，密码，角色，权限
    /// </summary>
    Task<LoginUser?> LoadUserByUsername(string username);

    /// <summary>
    /// 获取用户信息，根据前端需要返回
    /// </summary>
    virtual Task<Dictionary<string, object>> GetLoginUserInfo(LoginUser user)
    {
        var map = new Dictionary<string, object>
        {
            { "username", user.Username },
            { "roles", user.Roles },
            { "permissions", user.Permissions },
            { "data", user.UserData??"" },
        };
        if (!string.IsNullOrWhiteSpace(user.Nickname))
        {
            map.Add("nickname", user.Nickname);
        }
        return Task.FromResult(map);
    }

    /// <summary>
    /// 验证登录密码，默认直接通过密码明文对比,有加密的可重写该方法
    /// </summary>
    /// <param name="form">用户填写的登录表单数据</param>
    /// <param name="user">数据库中的用户信息</param>
    /// <returns>验证结果</returns>
    virtual bool ValidatePassword(LoginForm form, LoginUser user)
    {
        return form.Password == user.Password;
    }
}
```
#### 可实现ILoginService或继承LoginService重写相关方法来修改默认的登录接口实现逻辑
```
public interface ILoginService
{
    /// <summary>
    /// 登录系统
    /// </summary>
    Task<string> HandleLogin(LoginForm form);

    /// <summary>
    /// 退出登录
    /// </summary>
    /// <param name="token">要退出的token</param>
    /// <param name="user">当前操作用户</param>
    Task HandleLogout(string token, LoginUser user);

    /// <summary>
    /// 获取登录用户的详细信息，前端用啥，就返回啥，默认返回LoginUser里的信息
    /// </summary>
    Task<Dictionary<string, object>> GetLoginUserInfo(LoginUser user)
    {
        var map = new Dictionary<string, object>
        {
            { "username", user.Username },
            { "roles", user.Roles },
            { "permissions", user.Permissions },
        };
        return Task.FromResult(map);
    }

    /// <summary>
    /// 生成验证码
    /// </summary>
    /// <returns></returns>
    Task<CaptchaResult> CreateCaptcha()
    {
        return Task.FromResult(new CaptchaResult());
    }
}
```
### 2、认证模块
#### 采用token认证模块，系统默认实现Jwt和基于Cache的token
> 也可实现ITokenService接口，自定义token的实现方式，然后将实现类注入系统即可
```
public interface ITokenService
{
    /// <summary>
    /// 根据用户信息创建Token
    /// </summary>
    Task<string> CreateTokenAsync(LoginUser user);

    /// <summary>
    /// 验证Token合法性，通过后根据token获取用户信息,不通过写入原因
    /// </summary>
    Task<TokenValidationResult> ValidateTokenAsync(string token);

    /// <summary>
    /// 作废Token
    /// </summary>
    Task DisableTokenAsync(string token);
}
```

### 3、授权模块
 1. 超级管理员[LoginUser.IsSuperAdmin]用户拥用所有权限
 2. 提供特性[Permission]，可根据用户权限[LoginUser.Permissions]验证

## 四、日志模块

## 五、分页组件
使用 EFCore时，可以使用PageHelper实现分页
>该模块需要引入Microsoft.EntityFrameworkCore(>=8.0.0)


## 七、代码生成，使用EFCore Tools，反向工程    
  部分配置说明：efpt.renaming.json
  1. 去掉开头，添加后续的正则表达式配置
```
[
  {
    "SchemaName": "dbo",
    "UseSchemaName": false,
    "TableRegexPattern": "(^md|sys)_(?<table>.+$)",
    "TablePatternReplaceWith": "${table}_entity",
    "Tables": []
  }
]
```
  2. 需要注意sqlserver，需要SchemaName,mysql不能有SchemaName那一行
  
  3. 使用迁移说明Migration
        
    需要在WebApi项目添加 Microsoft.EntityFrameworkCore.Tools
    在程序包管理器控制台使用命令   Add-Migration，升级数据库使用 Update-Database
    如：  Add-Migration InitialCreate
    另外拆分项目时，需要配置 Assembly
    如：opt.UseSqlServer(conn, b => b.MigrationsAssembly("WebApi"));


## 常用扩展和工具类

## 系统所有配置项
以下配置值为系统默认值，只配置需要调整的部分即可。
 ```
{
  "FeiNuo": {
    "AppCode": "FeiNuo",        // 系统编码
    "AppName": "菲诺",          // 系统名称
    "AppVersion": "1.0.0",      // 版本号

    // 登录认证相关配置
    "Security": {
      "TokenType": "Jwt",      // Token类型：Jwt/Cache
      "TokenExpiration": 7200, // Token超时时间，单位秒，0不超时

      // Jwt配置
      "Jwt": {
        "SigningKey": "abcdefg1234", // 加密签名Key
        "ClockSkew": 1800,       // JWT的缓冲时间，单位秒，实际的过期时间 = exp+clockskew
        "Issuer": "",            // 证书颁发者，不配置默认不验证
        "Audience": ""           // 受众方，不配置默认不验证
      },

      // 验证码配置
      "Captcha": {
        "Enabled": true,                // 是否启用
        "Timeout": 300,                 // 超时时间,单位秒
        "Length": 4,                    // 验证码字符个数
        "Width": 100,                   // 图片宽度，单位px
        "Height": 40,                   // 图片高度，单位px
        "LineCount": 4,                 // 干扰线数量
        "ChaosDensity": 40,             // 燥点密度，长宽的像素积/密度计算出燥点个数，0表示不生成
        "FontSize": 30,                 // 字体大小中位数
        "BackGroundColor": "#ffffff",   // 图片背景色
      }
    }
  },
}
 ```
