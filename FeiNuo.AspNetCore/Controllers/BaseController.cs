using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.WebUtilities;

namespace FeiNuo.AspNetCore;

/// <summary>
/// Controller 基类
/// </summary>
[ApiController]
public class BaseController : ControllerBase
{
    #region 文件下载格式
    /// <summary>
    /// 文件下载格式:Excel
    /// </summary>
    protected const string CONTENT_TYPE_EXCEL = "application/vnd.ms-excel";
    /// <summary>
    /// 文件下载格式:WORD
    /// </summary>
    protected const string CONTENT_TYPE_WORD = "application/msword";
    /// <summary>
    /// 文件下载格式:PDF
    /// </summary>
    protected const string CONTENT_TYPE_PDF = "application/pdf";
    /// <summary>
    /// 文件下载格式:STREAM
    /// </summary>
    protected const string CONTENT_TYPE_STREAM = "application/octet-stream";
    #endregion

    #region 登录信息
    /// <summary>
    /// 当前登录用户对象
    /// </summary>
    protected LoginUser CurrentUser
    {
        get { return new LoginUser(User.Claims); }
    }

    /// <summary>
    /// 获取当前登录的token
    /// </summary>
    protected async Task<string?> GetAccessToken()
    {
        return await HttpContext.GetTokenAsync("access_token");
    }
    #endregion

    #region 返回信息
    protected static ActionResult InfoMessage(string message, object? data = null)
    {
        return Unprocessable(message, MessageTypeEnum.Info, data);
    }
    protected static ActionResult WarningMessage(string message, object? data = null)
    {
        return Unprocessable(message, MessageTypeEnum.Warning, data);
    }
    protected static ActionResult ErrorMessage(string message, object? data = null)
    {
        return Unprocessable(message, MessageTypeEnum.Error, data);
    }
    /// <summary>
    /// 返回422响应
    /// </summary>
    protected static ActionResult Unprocessable(string message, MessageTypeEnum type = MessageTypeEnum.Warning, object? data = null)
    {
        var resp = new MessageResult(message, type, data);
        return new UnprocessableEntityObjectResult(resp);
    }
    #endregion

    protected Dictionary<string, string> GetRequestParam()
    {
        var paramMap = QueryHelpers.ParseQuery(Request.QueryString.Value);
        var dic = new Dictionary<string, string>();
        foreach (var param in paramMap)
        {
            dic.Add(param.Key, param.Value.ToString());
        }
        return dic;
    }
}