# FeiNuo.AspNetCore
## 功能介绍
    本项目基于Net8开发，用于辅助开发WebApi项目。    
1. 封装自动注入，异常处理，登录认证，系统授权等通用代码；
2. 提供常用工具类和扩展功能：如JsonUtil,StringExtensions等;
3. 提供常用组件的封装，如：Excel操作，日志操作等,验证码生成等；

## 一、自动注入
    builder.Services.AddAppServices();
1. 系统自动注入实现IService接口或继承BaseService类的所有类
    1. 如果该类有实现的接口(除IService)则根据接口类注入
    2. 如果没有实现接口(除IService)，则根据当前类型注入
    3. 注入的生命周期全部为 `ServiceLifetime.Scoped`
    4. 如果默认注入的不合适，可添加[Service]特性自定义
    
2. 系统自动注入标注有[Service]特性的所有类
    1. 如果特性中有注入类型，则以特性中指定的类型为准
    2. 如果特性中没有注入类型，则根据当前类实现的接口注入
    3. 如果当前类也没有实现接口，则根据当前类型注入
    4. 生命周期默认为 `ServiceLifetime.Scoped`，可以在特性中修改
    
3.  默认注入内存缓存和内存分布式缓存

## 二、API行为处理
    builder.Services.AddAppControllers();
1. 系统约定统一根据HTTP状态码返回数据
2. 支持DateOnly类型接收带时间的日期
3. 正常响应，返回状态码2XX,直接返回系统数据
4. 异常响应：
    1. 没有登录信息：返回 401,UnauthorizedResult
    2. 没有权限：返回 403,ForbidResult
    3. 模型数据验证不通过：返回 400,BadRequestResult
    4. 系统自定义异常消息：返回 422,UnprocessableEntityResult
    5. 系统异常：返回 500
1. 统一序列化配置
    1. null值默认不输出
    1. 阻断循环引用
    1. 属性转换策略：小驼峰
    1. 支持从字符串转成数字
    1. 日期类型默认输出格式 yyyy-MM-dd HH:mm:ss
    
## 三、认证授权
    builder.Services.AddAppSecurity(builder.Configuration);

### 1、登录接口
    提供登录相关Rest接口  
    需要需要注入查询登录用户的实现类：ILoginUserService  

|    接口   | 类型  |     参数   |   返回  |
|    ----   | :---: |    ------  |   ----  |
| /login    | POST  | LoginForm  | Token字体为串
| /logout   | POST  | 无         | 无
| /userinfo | GET   | 无         | LoginUser
| /captcha  | GET   | 无         | CaptchaResult  

#### 必需实现 ILoginService，以提供登录所需的用户信息
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
#### 可实现ILoginService或继承LoginService重新相关方法来修改默认的登录接口实现逻辑
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
/// <summary>
    /// Token服务类
    /// </summary>
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

## 五分页组件

## PoiExcel封装

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
