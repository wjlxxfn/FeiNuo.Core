using Microsoft.Extensions.DependencyInjection;

namespace FeiNuo.AspNetCore;

/// <summary>
/// 标识为服务类的注解，添加注解后可自动注入
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class ServiceAttribute : Attribute
{
    public ServiceAttribute(ServiceLifetime lifetime = ServiceLifetime.Scoped, params Type[] serviceTypes)
    {
        Lifetime = lifetime;
        ServiceTypes = serviceTypes;
    }

    /// <summary>
    /// 生命周期
    /// </summary>
    public ServiceLifetime Lifetime { get; set; }

    /// <summary>
    /// 注入的接口类型
    /// </summary>
    public Type[] ServiceTypes { get; set; }
}