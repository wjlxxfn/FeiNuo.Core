using System.ComponentModel;

namespace FeiNuo.Core;

/// <summary>
/// 请求的客户端信息
/// </summary>
public class RequestClient
{
    /// <summary>
    /// 客户端IP
    /// </summary>
    [Description("客户端IP")]
    public string ClientIp { get; set; } = string.Empty;

    /// <summary>
    /// 客户端操作系统
    /// </summary>
    [Description("客户端操作系统")]
    public string ClientOs { get; set; } = string.Empty;

    /// <summary>
    /// 客户端浏览器
    /// </summary>
    [Description("客户端浏览器")]
    public string ClientBrowser { get; set; } = string.Empty;

    /// <summary>
    /// 客户端设备
    /// </summary>
    [Description("客户端设备")]
    public string ClientDevice { get; set; } = string.Empty;

    /// <summary>
    /// 是否移动端
    /// </summary>
    [Description("是否移动端")]
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
