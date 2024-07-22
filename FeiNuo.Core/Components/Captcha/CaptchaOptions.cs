namespace FeiNuo.Core.Captcha
{
    /// <summary>
    /// 验证码配置项
    /// </summary>
    public class CaptchaOptions
    {
        /// <summary>
        /// 验证码字符个数
        /// </summary>
        public int Length { get; set; } = 4;

        /// <summary>
        /// 图片宽度，单位px
        /// </summary>
        public int Width { get; set; } = 100;

        /// <summary>
        /// 图片高度，单位px
        /// </summary>
        public int Height { get; set; } = 40;

        /// <summary>
        /// 干扰线数量,默认4条
        /// </summary>
        public int LineCount { get; set; } = 4;

        /// <summary>
        /// 燥点密度，长宽的像素积/密度计算出燥点个数，0表示不生成
        /// </summary>
        public int ChaosDensity { get; set; } = 40;

        /// <summary>
        /// 字体大小中位数，默认30
        /// </summary>
        public int FontSize { get; set; } = 30;

        /// <summary>
        /// 图片背景色，默认白色 #ffffff
        /// </summary>
        public string BackGroundColor { get; set; } = "#ffffff";

        /// <summary>
        /// 生成的图片类型，默认gif，可选值：gif,png,jpg,jpeg,bmp
        /// </summary>
        public string ImageType { get; set; } = "gif";
    }
}
