using FeiNuo.Core.Utilities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.DependencyInjection;
using System.Diagnostics;

namespace FeiNuo.AspNetCore;

/// <summary>
/// 操作日志记录属性
/// </summary>
[AttributeUsage(AttributeTargets.Method)]
public class LogAttribute : ActionFilterAttribute
{
    private readonly string _logTitle;
    private readonly OperateType _operType;
    private readonly bool _saveParam;
    private readonly bool _saveResult;
    public LogAttribute(string logTitle, OperateType operType, bool saveParam = true, bool saveResult = false)
    {
        _logTitle = logTitle;
        _operType = operType;
        _saveParam = saveParam;
        _saveResult = saveResult;
    }
    /// <summary>
    /// 记录操作日志
    /// </summary>
    public override async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
    {
        // 执行action，记录执行耗时
        var stopwatch = Stopwatch.StartNew();
        var resultContext = await next();

        // 生成操作日志
        OperateLog log = CreateOperateLog(context, resultContext);

        // 记录执行时间
        stopwatch.Stop();
        log.ExecuteTime = stopwatch.ElapsedMilliseconds;

        // 保存日志
        var logService = resultContext.HttpContext.RequestServices.GetRequiredService<ILogService>();
        await logService.SaveLog(log);
    }

    private OperateLog CreateOperateLog(ActionExecutingContext context, ActionExecutedContext resultContext)
    {
        // string controller = context.Controller.ToString() ?? "";
        // string action = context.ActionDescriptor.DisplayName ?? "";

        var log = new OperateLog(_operType, _logTitle);
        log.MergeContextParam(resultContext.HttpContext);
        if (_saveParam && context.ActionArguments.Count > 0)
        {
            if (context.ActionArguments.Count == 1)
            {
                var p = context.ActionArguments.Values.First();
                if (p != null && (p is string || p.GetType().IsPrimitive))
                {
                    log.RequestParam = p.ToString() ?? "";
                }
            }
            else
            {
                log.RequestParam = JsonUtils.Serialize(context.ActionArguments);
            }
        }
        // 有异常说明失败
        log.Success = resultContext.Exception == null;
        if (!log.Success)
        {
            log.LogContent = $"操作异常：";
            // DbUpdateException
            if (resultContext.Exception!.InnerException != null)
            {
                log.LogContent += resultContext.Exception.InnerException.Message;
            }
            else
            {
                log.LogContent += resultContext.Exception!.Message;
            }
        }
        else if (_saveResult && resultContext.Result != null)
        {
            if (resultContext.Result is ObjectResult obj)
            {
                if (obj.Value == null || obj.Value.ToString() == "")
                {
                    log.LogContent = "";
                }
                else
                {
                    if (obj.Value is string || obj.Value.GetType().IsPrimitive)
                    {
                        log.LogContent = obj.Value.ToString() ?? "";
                    }
                    else
                    {
                        log.LogContent = JsonUtils.Serialize(obj.Value);
                    }
                }
            }
            else if (resultContext.Result is ContentResult cr)
            {
                log.LogContent = cr.Content ?? "";
            }
            else if (resultContext.Result is EmptyResult)
            {
                log.LogContent = "";
            }
            else
            {
                log.LogContent = JsonUtils.Serialize(resultContext.Result);
            }
        }

        return log;
    }
}