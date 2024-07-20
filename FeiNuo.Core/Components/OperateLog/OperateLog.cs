namespace FeiNuo.Core
{
    /// <summary>
    /// 操作日志对象
    /// </summary>
    public class OperateLog
    {
        /// <summary>
        /// 默认无参构造函数
        /// </summary>
        public OperateLog() { }

        /// <summary>
        /// 构造函数
        /// </summary>
        public OperateLog(OperateType operType, string logTitle, string logContent = "")
        {
            LogTitle = logTitle;
            OperateType = operType;
            LogContent = logContent;
            OperateTime = DateTime.Now;
            Success = true;
        }

        /// <summary>
        /// 操作类型
        /// </summary>
        public OperateType OperateType { get; set; }

        /// <summary>
        /// 日志标题
        /// </summary>
        public string LogTitle { get; set; } = string.Empty;

        /// <summary>
        /// 日志内容，或响应内容
        /// </summary>
        public string LogContent { get; set; } = string.Empty;

        /// <summary>
        /// 请求URL
        /// </summary>
        public string RequestPath { get; set; } = string.Empty;

        /// <summary>
        /// 请求方法
        /// </summary>
        public string RequestMethod { get; set; } = string.Empty;

        /// <summary>
        /// 请求参数
        /// </summary>
        public string RequestParam { get; set; } = string.Empty;

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
        /// 是否执行成功
        /// </summary>
        public bool Success { get; set; } = true;

        /// <summary>
        /// 执行时长,毫秒
        /// </summary>
        public long ExecuteTime { get; set; }

        /// <summary>
        /// 操作人
        /// </summary>
        public string OperateBy { get; set; } = string.Empty;

        /// <summary>
        /// 操作时间
        /// </summary>
        public DateTime OperateTime { get; set; } = DateTime.Now;

    }
}
