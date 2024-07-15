namespace FeiNuo.Core
{
    /// <summary>
    /// 请求的客户端信息
    /// </summary>
    public class RequestClient
    {
        /// <summary>
        /// 客户端IP
        /// </summary>
        public string ClientIp { get; set; } = string.Empty;

        /// <summary>
        /// 客户端操作系统
        /// </summary>
        public string ClientOs { get; set; } = string.Empty;

        /// <summary>
        /// 客户端浏览器
        /// </summary>
        public string ClientBrowser { get; set; } = string.Empty;

        /// <summary>
        /// 客户端设备
        /// </summary>
        public string ClientDevice { get; set; } = string.Empty;

        /// <summary>
        /// 是否移动端
        /// </summary>
        public bool IsMobile { get; set; }

        /// <summary>
        /// 默认无参构造函数
        /// </summary>
        public RequestClient()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        public RequestClient(string clientIp, string clientOs, string clientBrowser, string clientDevice, bool isMobile)
        {
            ClientIp = clientIp;
            ClientOs = clientOs;
            ClientBrowser = clientBrowser;
            ClientDevice = clientDevice;
            IsMobile = isMobile;
        }
    }
}
