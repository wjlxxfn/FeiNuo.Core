using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.Net;

namespace FeiNuo.Core
{

    /// <summary>
    /// 自定义异常，拦截掉，不让进入ExceptionHandlerMiddleware中间件
    /// 本来可以都在/error中处理，但是进入/error之前会经过ExceptionHandlerMiddleware中间件，里面又会打印error日志
    /// 另外http请求也会变成两个。httplog日志会多显示一条，通过该过滤器可以直接修改响应内容
    /// </summary>
    public class AppExceptionFilter(IWebHostEnvironment hostEnv, ILogger<AppExceptionFilter> logger) : IAsyncExceptionFilter
    {
        private readonly IWebHostEnvironment _hostEnv = hostEnv;
        private readonly ILogger<AppExceptionFilter> _logger = logger;

        /// <summary>
        /// 出异常后判断如果是MsgException，拦截掉，并标识异常已处理
        /// </summary>
        public Task OnExceptionAsync(ExceptionContext context)
        {
            if (!context.ExceptionHandled)
            {
                var exception = context.Exception;

                var msgVO = new MessageResult(exception.Message, MessageType.Error);
                var httpStatus = HttpStatusCode.InternalServerError;

                // 自定义消息异常，不需要记录日志，返回422
                if (exception is MessageException msgExp)
                {
                    msgVO.Type = msgExp.MessageType;
                    httpStatus = HttpStatusCode.UnprocessableEntity;
                }
                else
                {
                    _logger.LogError(exception, "系统异常：{Message}", exception.Message);
                    // EF具体异常记录在InnerException
                    if (exception is DbUpdateException)
                    {
                        exception = exception.InnerException!;
                        msgVO.Message = exception.Message;
                    }
                    // 开发环境显示日志详细信息
                    if (_hostEnv.IsDevelopment())
                    {
                        msgVO.Data = exception.StackTrace;
                    }
                }
                // 返回MessageVO
                context.Result = new ObjectResult(msgVO)
                {
                    StatusCode = (int)httpStatus
                };
                // 标识为已处理
                context.ExceptionHandled = true;
            }
            return Task.CompletedTask;
        }
    }
}