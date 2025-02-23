using Microsoft.Extensions.Logging;

namespace FeiNuo.AspNetCore;

/// <summary>
/// 默认操作日志记录服务，注入系统的日志记录服务
/// 日志级别Information，可通过日志配置输出到数据库，
/// 也可新写日志记录类实现ILogService接口并注入框架
/// </summary>
internal class SimpleLogService : ILogService
{
    private readonly ILogger<SimpleLogService> _logger;

    public SimpleLogService(ILoggerFactory logFactory)
    {
        _logger = logFactory.CreateLogger<SimpleLogService>();
        _logger.LogWarning("未注入日志服务，使用默认的日志组件ILogger输出Info类型的日志.");
    }

    public async Task SaveLog(OperateLog log)
    {
        await Task.Run(() =>
        {
            var Result = log.Success ? "成功" : "失败";
            _logger.LogInformation("【操作日志】【{RequestMethod}】{RequestPath} , {LogTitle}: 执行{Result} ,耗时{ExecuteTime} ms, 用户：{Operator} \n详情：{@LogDetail}",
                log.RequestMethod, log.RequestPath, log.LogTitle, Result, log.ExecuteTime, log.OperateBy, log);
        });
    }

    public async Task SaveLog(OperateType operateType, string logTitle, string logDetail, RequestClient? request = null)
    {
        var log = new OperateLog(operateType, logTitle, logDetail);
        if (request != null)
        {
            log.MergeClientInfo(request);
        }
        await SaveLog(log);
    }
}
