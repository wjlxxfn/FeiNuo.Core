﻿using FeiNuo.Core.Utilities;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;

namespace FeiNuo.Core
{
    /// <summary>
    /// 默认操作日志记录服务，注入系统的日志记录服务
    /// 日志级别Information，可通过日志配置输出到数据库，
    /// 也可新写日志记录类实现ILogService接口并注入框架
    /// </summary>
    internal class SimpleLogService : ILogService
    {
        private readonly ILogger<SimpleLogService> _logger;
        private readonly IHttpContextAccessor _httpContextAccessor;

        public SimpleLogService(ILoggerFactory logFactory, IHttpContextAccessor httpContextAccessor)
        {
            _httpContextAccessor = httpContextAccessor;
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

        public async Task SaveLog(OperateType operateType, string logTitle, string logDetail)
        {
            var log = new OperateLog(operateType, logTitle, logDetail);
            log.MergeContextParam(_httpContextAccessor.HttpContext);
            await SaveLog(log);
        }
    }

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
            log.ClientIp = client.ClientIp;
            log.ClientOs = client.ClientOs;
            log.ClientBrowser = client.ClientBrowser;
            log.ClientDevice = client.ClientDevice;
            log.IsMobile = client.IsMobile;
        }
    }
}
