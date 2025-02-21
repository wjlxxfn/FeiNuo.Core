namespace FeiNuo.AspNetCore.Security;

public class SecurityOptions
{
    /// <summary>
    /// appsettings.json中的配置key
    /// </summary>
    public const string ConfigKey = AppConfig.ConfigKey + ":Security";

    /// <summary>
    /// Token超时时间，单位秒，0不超时，默认2小时
    /// </summary>
    public int TokenExpiration { get; set; } = 7200;

    /// <summary>
    /// Jwt参数配置
    /// </summary>
    public JwtOptions Jwt { get; set; } = new JwtOptions();

    /// <summary>
    /// 验证码参数配置
    /// </summary>
    public CaptchaOptions Captcha { get; set; } = new CaptchaOptions();
}


#region jwt 配置
/// <summary>
/// jwt 配置参数
/// </summary>
public class JwtOptions
{
    /// <summary>
    /// appsettings.json中的配置key
    /// </summary>
    public const string ConfigKey = SecurityOptions.ConfigKey + ":Jwt";

    /// <summary>
    /// 加密签名Key
    /// </summary>
    public string SigningKey { get; set; } = "pt&wjl[6xxfn8]%.";

    /// <summary>
    /// JWT的缓冲时间，默认30分钟
    /// </summary>
    /// <remarks>
    /// 实际的过期时间 = exp+clockskew <br/>
    /// 系统使用该时间实现jwt的滑动过期刷新token，为0的话则不会自动刷新token <br/>
    /// 比如clockskew=30分钟，则在到期时间前30分钟或后30分钟访问时都会重新生成token <br/>
    /// 在前30分钟刷新token有个前提，过期时间必须大于缓冲时间的两倍，避免太频繁刷新
    /// </remarks>
    public int ClockSkew { get; set; } = 1800;

    /// <summary>
    /// 证书颁发者，不配置默认不验证
    /// </summary>
    public string? Issuer { get; set; }

    /// <summary>
    /// 受众方，不配置默认不验证
    /// </summary>
    public string? Audience { get; set; }

    /// <summary>
    /// 如果设置为true, 退出登录后，jwt加入到黑名单的缓存中，每次访问会判断黑名单，会影响性能，安全要求较高时在配置
    /// </summary>
    public bool CheckForbidden { get; set; } = false;
}
#endregion

#region 验证码
public class CaptchaOptions : Core.Captcha.CaptchaOptions
{
    /// <summary>
    /// appsettings.json中的配置key
    /// </summary>
    public const string ConfigKey = SecurityOptions.ConfigKey + ":Captcha";

    /// <summary>
    /// 是否启用
    /// </summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    /// 超时时间,单位秒，默认 5 分钟
    /// </summary>
    public int Timeout { get; set; } = 300;
}
#endregion
