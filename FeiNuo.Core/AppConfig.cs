namespace FeiNuo.Core
{
    /// <summary>
    /// 系统配置项
    /// </summary>
    public class AppConfig
    {
        /// <summary>
        /// appsettings.json中的配置key
        /// </summary>
        public const string ConfigKey = "FeiNuo";

        /// <summary>
        /// 系统编码
        /// </summary>
        public string AppCode { get; set; } = "FeiNuo";

        /// <summary>
        /// 系统名称
        /// </summary>
        public string AppName { get; set; } = "菲诺";

        /// <summary>
        /// 版本号
        /// </summary>
        public string AppVersion { get; set; } = "1.0.0";

    }
}
