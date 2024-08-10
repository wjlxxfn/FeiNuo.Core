using Microsoft.Extensions.DependencyInjection;
using System.Reflection;

namespace FeiNuo.Core
{
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

    public static class ServiceInjectionExtensions
    {
        public static IServiceCollection AutoInjectServcice(this IServiceCollection services)
        {
            // 扫描所有类
            var types = AppDomain.CurrentDomain.GetAssemblies().SelectMany(x => x.GetTypes());

            #region 自动注入服务类：实现[IBaseServicer]
            var businessServiceTypes = types.Where(t => typeof(IService).IsAssignableFrom(t) && t.IsClass && !t.IsAbstract);
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
            var attributeTypes = types.Where(t => t.GetCustomAttributes(typeof(ServiceAttribute), false).Length > 0 && t.IsClass && !t.IsAbstract);
            foreach (var type in attributeTypes)
            {
                var attr = type.GetCustomAttribute<ServiceAttribute>()!;
                var lifetime = attr.Lifetime;
                // 有指定接口类型的，根据接口类型注入
                if (null != attr.ServiceTypes && attr.ServiceTypes.Length > 0)
                {
                    attr.ServiceTypes.ToList().ForEach(t => AddService(services, lifetime, t, type));
                }
                else
                {
                    // 有指定接口的注入接口
                    var interfaces = type.GetInterfaces().ToList();
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
}