using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace FeiNuo.AspNetCore.Security.FormLogin;

/// <summary>
/// 登录表单
/// </summary>
public class LoginForm
{
    /// <summary>
    /// 用户名
    /// </summary>
    [Description("用户名")]
    [Required(ErrorMessage = "用户名必填")]
    public string Username { get; set; } = string.Empty;

    /// <summary>
    /// 密码
    /// </summary>
    [Description("密码")]
    [Required(ErrorMessage = "密码必填")]
    public string Password { get; set; } = string.Empty;

    /// <summary>
    ///  验证码值，
    /// </summary>
    [Description("验证码值")]
    public string Captcha { get; set; } = string.Empty;

    /// <summary>
    ///  验证码需要一个存储的key值，前端需要一起传过来
    /// </summary>
    [Description("验证码键")]
    public string CaptchaKey { get; set; } = string.Empty;

}