using Microsoft.Extensions.Logging;

namespace FeiNuo.Core;

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

    public Task SaveLog(OperateLog log)
    {
        _ = Task.Run(() =>
        {
            var Result = log.Success ? "成功" : "失败";
            _logger.LogInformation("【操作日志】【{RequestMethod}】{RequestPath} , {LogTitle}: 执行{Result} ,耗时{ExecuteTime} ms, 用户：{Operator} \n详情：{@LogDetail}",
                log.RequestMethod, log.RequestPath, log.LogTitle, Result, log.ExecuteTime, log.OperateBy, log);
        });
        return Task.CompletedTask;
    }

    public async Task SaveLog(OperateType operateType, string logTitle, string logDetail, RequestClient? client = null)
    {
        var log = new OperateLog(operateType, logTitle, logDetail);
        if (client != null)
        {
            log.RequestClient = client;
        }
        await SaveLog(log);
    }
}
