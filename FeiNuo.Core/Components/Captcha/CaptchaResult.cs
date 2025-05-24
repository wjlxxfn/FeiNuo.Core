using System.ComponentModel;

namespace FeiNuo.Core.Captcha;

public class CaptchaResult
{
    private CaptchaResult() { }

    /// <summary>
    /// 使用该对象表示禁用验证码
    /// </summary>
    public static CaptchaResult Disabled => new CaptchaResult() { Enabled = false };

    public CaptchaResult(string captchaKey, string captcha)
    {
        Enabled = true;
        Captcha = captcha;
        CaptchaKey = captchaKey;
    }

    /// <summary>
    /// 是否启用验证码
    /// </summary>
    [Description("是否启用验证码")]
    public bool Enabled { get; private set; } = true;

    /// <summary>
    /// 验证码内容
    /// </summary>
    [Description("验证码内容")]
    public string CaptchaKey { get; private set; } = string.Empty;

    /// <summary>
    /// 图片的base64编码
    /// </summary>
    [Description("图片的base64编码")]
    public string Captcha { get; private set; } = string.Empty;
}
