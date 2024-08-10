namespace FeiNuo.Core.Security
{
    public class CaptchaResult
    {
        public CaptchaResult()
        {
            Enabled = false;
        }

        public CaptchaResult(string captchaKey, string captcha)
        {
            Enabled = true;
            Captcha = captcha;
            CaptchaKey = captchaKey;
        }

        public bool Enabled { get; private set; } = true;

        public string CaptchaKey { get; private set; } = string.Empty;

        /// <summary>
        /// 图片的base64编码
        /// </summary>
        public string Captcha { get; private set; } = string.Empty;
    }
}
