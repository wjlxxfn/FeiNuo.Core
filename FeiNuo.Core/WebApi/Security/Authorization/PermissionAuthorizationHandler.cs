using Microsoft.AspNetCore.Authorization;

namespace FeiNuo.Core.Security;

/// <summary>
/// 鉴权处理类
/// </summary>
internal class PermissionAuthorizationHandler : AuthorizationHandler<PermissionAttribute>
{
    /// <summary>
    /// 鉴权：根据权限字符串鉴权
    /// </summary>
    protected override Task HandleRequirementAsync(AuthorizationHandlerContext context, PermissionAttribute requirement)
    {
        if (context.User != null)
        {
            var user = new LoginUser(context.User.Claims);
            // 有其中一个权限即可
            if (user.Permissions.Any(t => requirement.Permission.Split(',').Contains(t)))
            {
                context.Succeed(requirement);
            }
        }
        return Task.CompletedTask;
    }
}
