namespace FeiNuo.Core.Captcha
{
    public class CaptchaData
    {
        /// <summary>
        /// 验证码内容
        /// </summary>
        public string Text { get; set; } = string.Empty;

        /// <summary>
        /// 验证码图片字节数组
        /// </summary>
        public byte[] ImageBytes { get; set; } = null!;

        /// <summary>
        /// 图片格式，默认gif，可选值：gif,png,jpg,jpeg,bmp
        /// </summary>
        public string ImageFormat { get; set; } = null!;

        /// <summary>
        /// 图片转成base64
        /// </summary>
        public string ImageBase64
        {
            get
            {
                return Convert.ToBase64String(ImageBytes);
            }
        }
    }
}
