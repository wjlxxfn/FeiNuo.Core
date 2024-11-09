namespace FeiNuo.AspNetCore.Security;

public class SecurityOptions
{
    /// <summary>
    /// appsettings.json中的配置key
    /// </summary>
    public const string ConfigKey = AppConfig.ConfigKey + ":Security";

    /// <summary>
    /// Token类型：Jwt,Cache或Other
    /// </summary>
    public string TokenType { get; set; } = "Jwt";

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
    /// JWT的缓冲时间，实际的过期时间 = exp+clockskew<br/>
    /// 系统使用该时间实现token的滑动过期，即token过期但在缓冲时间内会生成刷新token<br/>
    /// </summary>
    public int ClockSkew { get; set; } = 1800;

    /// <summary>
    /// 证书颁发者，不配置默认不验证
    /// </summary>
    public string? Issuer { get; set; }

    /// <summary>
    /// 受众方，不配置默认不验证
    /// </summary>
    public string? Audience { get; set; }
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
