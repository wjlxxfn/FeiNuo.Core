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
        log.RequestClient = ClientUtils.GetClientInfo(context);
    }
}
