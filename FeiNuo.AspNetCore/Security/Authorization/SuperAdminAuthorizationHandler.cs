using Microsoft.AspNetCore.Authorization;

namespace FeiNuo.Core.Security
{
    /// <summary>
    /// 超级管理员拥有所有权限
    /// </summary>
    internal class SuperAdminAuthorizationHandler : IAuthorizationHandler
    {
        public Task HandleAsync(AuthorizationHandlerContext context)
        {
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
}
