using System.ComponentModel.DataAnnotations;

namespace FeiNuo.Core;

/// <summary>
/// 登录表单
/// </summary>
public class LoginForm
{
    /// <summary>
    /// 用户名
    /// </summary>
    [Required(ErrorMessage = "用户名必填")]
    public string Username { get; set; } = string.Empty;

    /// <summary>
    /// 密码
    /// </summary>
    [Required(ErrorMessage = "密码必填")]
    public string Password { get; set; } = string.Empty;

    /// <summary>
    ///  验证码值，
    /// </summary>
    public string Captcha { get; set; } = string.Empty;

    /// <summary>
    ///  验证码需要一个存储的key值，前端需要一起传过来
    /// </summary>
    public string CaptchaKey { get; set; } = string.Empty;

}