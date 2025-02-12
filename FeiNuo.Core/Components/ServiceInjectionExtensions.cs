﻿using Microsoft.Extensions.DependencyInjection;
using System.Reflection;

namespace FeiNuo.Core;

public static class ServiceInjectionExtensions
{
    public static IServiceCollection AutoInjectServcice(this IServiceCollection services)
    {
        // 扫描所有类
        var types = AppDomain.CurrentDomain.GetAssemblies()
            .SelectMany(x => x.GetTypes())
            .Where(t => t.IsClass && !t.IsAbstract);

        #region 自动注入服务类：实现[IServicer]，排除掉Service特性，有特性的以特性为准
        var businessServiceTypes = types.Where(t =>
                typeof(IService).IsAssignableFrom(t)
                && t.GetCustomAttributes(typeof(ServiceAttribute), false).Length == 0
            );
        foreach (var type in businessServiceTypes)
        {
            // 服务类有接口的根据接口注入
            var interfaces = type.GetInterfaces().Where(a => a != typeof(IService));
            if (interfaces.Any())
            {
                interfaces.ToList().ForEach(r => AddService(services, ServiceLifetime.Scoped, r, type));
            }
            // 没有接口的直接根据当前类的类型注入
            else AddService(services, ServiceLifetime.Scoped, type, type);
        }
        #endregion

        #region 自动注入服务类：注解[ServiceAttribute]
        var attributeTypes = types.Where(t => t.GetCustomAttributes(typeof(ServiceAttribute), false).Length > 0);
        foreach (var type in attributeTypes)
        {
            var attr = type.GetCustomAttribute<ServiceAttribute>()!;
            var lifetime = attr.Lifetime;
            // 在特性中有指定接口类型的，根据特性接口类型注入
            if (null != attr.ServiceTypes && attr.ServiceTypes.Length > 0)
            {
                attr.ServiceTypes.ToList().ForEach(t => AddService(services, lifetime, t, type));
            }
            else
            {
                // 有实现接口的注入接口
                var interfaces = type.GetInterfaces().Where(a => a != typeof(IService) && a != typeof(BaseService)).ToList();
                if (interfaces.Count > 0)
                {
                    foreach (var r in interfaces)
                    {
                        AddService(services, ServiceLifetime.Scoped, r, type);
                    }
                }
                else // 没有在注解中指定接口类型,也没有实现接口，则直接注入原类型
                {
                    AddService(services, lifetime, type, type);
                }
            }
        }
        #endregion

        return services;
    }

    /// <summary>
    /// 根据生命周期注入服务类
    /// </summary>
    private static void AddService(IServiceCollection services, ServiceLifetime lifetime, Type serviceType, Type implementationType)
    {
        switch (lifetime)
        {
            case ServiceLifetime.Singleton:
                services.AddSingleton(serviceType, implementationType);
                break;
            case ServiceLifetime.Scoped:
                services.AddScoped(serviceType, implementationType);
                break;
            case ServiceLifetime.Transient:
                services.AddTransient(serviceType, implementationType);
                break;
        }
    }
}
