using Microsoft.AspNetCore.Authorization;

namespace FeiNuo.AspNetCore;

/// <summary>
/// 通过字符串实现权限控制
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = true, Inherited = true)]
public class PermissionAttribute(string permission) : AuthorizeAttribute, IAuthorizationRequirement, IAuthorizationRequirementData
{
    /// <summary>
    /// 所需的权限字符串，多个用逗号分隔，只要有一个满足就通过
    /// </summary>
    public string Permission { get; set; } = permission;

    public IEnumerable<IAuthorizationRequirement> GetRequirements()
    {
        yield return this;
    }
}
