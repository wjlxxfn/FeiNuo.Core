using FeiNuo.AspNetCore.Utilities;
using Microsoft.AspNetCore.Http;

namespace FeiNuo.AspNetCore;

public static class OperateLogExtensions
{
    /// <summary>
    /// 添加Reqest相关参数
    /// </summary>
    public static void MergeContextParam(this OperateLog log, HttpContext? context)
    {
        if (null == context) return;
        // User
        log.OperateBy = context.User.Identity?.Name ?? "";

        // Request
        var request = context.Request;
        log.RequestMethod = request.Method;
        log.RequestPath = request.Path;
        log.RequestParam = request.QueryString.ToString();

        // Client
        var client = ClientUtils.GetClientInfo(context);
        log.MergeClientInfo(client);
    }

    /// <summary>
    /// 添加Reqest相关参数
    /// </summary>
    public static void MergeClientInfo(this OperateLog log, RequestClient? client)
    {
        if (null == client) return;

        log.ClientIp = client.ClientIp;
        log.ClientOs = client.ClientOs;
        log.ClientBrowser = client.ClientBrowser;
        log.ClientDevice = client.ClientDevice;
        log.IsMobile = client.IsMobile;
    }
}
