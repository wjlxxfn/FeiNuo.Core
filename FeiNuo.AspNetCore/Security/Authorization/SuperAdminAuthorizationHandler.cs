using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Logging;

namespace FeiNuo.AspNetCore.Security.Authorization;

/// <summary>
/// 超级管理员拥有所有权限
/// </summary>
public class SuperAdminAuthorizationHandler : IAuthorizationHandler
{
    private readonly ILogger<SuperAdminAuthorizationHandler> logger;

    public SuperAdminAuthorizationHandler(ILogger<SuperAdminAuthorizationHandler> logger)
    {
        this.logger = logger;
    }

    public Task HandleAsync(AuthorizationHandlerContext context)
    {
        if (logger.IsEnabled(LogLevel.Debug))
        {
            logger.LogDebug("CheckSuperAdmin,User = {User}", context.User.Identity?.Name);
        }

        if (context.User != null && (context.User.Identity?.IsAuthenticated ?? false))
        {
            var user = new LoginUser(context.User.Claims);
            if (user.IsSuperAdmin)
            {
                context.PendingRequirements.ToList().ForEach(context.Succeed);
            }
        }
        return Task.CompletedTask;
    }
}
