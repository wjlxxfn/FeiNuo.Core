using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Logging;

namespace FeiNuo.AspNetCore.Security.Authorization;

/// <summary>
/// 鉴权处理类
/// </summary>
public class PermissionAuthorizationHandler : AuthorizationHandler<PermissionAttribute>
{
    private readonly ILogger<PermissionAuthorizationHandler> logger;

    public PermissionAuthorizationHandler(ILogger<PermissionAuthorizationHandler> logger)
    {
        this.logger = logger;
    }
    /// <summary>
    /// 鉴权：根据权限字符串鉴权
    /// </summary>
    protected override Task HandleRequirementAsync(AuthorizationHandlerContext context, PermissionAttribute requirement)
    {
        if (logger.IsEnabled(LogLevel.Debug))
        {
            logger.LogDebug("CheckPermission:User = {User},RequirePermissions = {Permission}", context.User.Identity?.Name ?? "", requirement.Permission);
        }
        if (context.User.Identity != null)
        {
            var user = new LoginUser(context.User.Claims);
            var perms = requirement.Permission.Split(',');

            // 有其中一个权限即可
            if (perms.Any(user.Permissions.Contains))
            {
                context.Succeed(requirement);
            }
        }
        return Task.CompletedTask;
    }
}
