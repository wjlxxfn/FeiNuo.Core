using FeiNuo.Core.Utilities;
using Microsoft.Extensions.Caching.Distributed;

namespace FeiNuo.Core;

/// <summary>
/// 分布式缓存扩展方法
/// </summary>
public static class DistributedCacheExtensions
{
    /// <summary>
    /// 设置缓存
    /// </summary>
    /// <param name="cache">IDistributedCache</param>
    /// <param name="key">缓存键</param>
    /// <param name="value">缓存对象</param>
    /// <param name="options">缓存选项：不传默认为滑动过期24小时</param>
    public static Task SetObjectAsync<T>(this IDistributedCache cache, string key, T value, DistributedCacheEntryOptions? options = null) where T : class
    {
        if (value == null || value.ToString() == null)
        {
            throw new ArgumentNullException(nameof(value));
        }
        options ??= new DistributedCacheEntryOptions() { SlidingExpiration = TimeSpan.FromHours(24) };
        var cacheValue = value is string strValue ? strValue : JsonUtils.Serialize(value);
        return cache.SetStringAsync(key, cacheValue, options);
    }

    /// <summary>
    /// 获取对象类型的缓存值
    /// </summary>
    public static async Task<T?> GetObjectAsync<T>(this IDistributedCache cache, string key) where T : class
    {
        var data = await cache.GetStringAsync(key);
        if (data == null)
        {
            return null;
        }
        return JsonUtils.Deserialize<T>(data);
    }

    /// <summary>
    /// 获取对象类型的缓存值，没有缓存时，查询数据后自动添加缓存
    /// </summary>
    /// <param name="cache"></param>
    /// <param name="key"></param>
    /// <param name="factory"></param>
    /// <param name="options">缓存选项：不传默认为滑动过期24小时</param>
    public static async Task<T?> GetOrCreateAsync<T>(this IDistributedCache cache, string key, Func<DistributedCacheEntryOptions, T> factory, DistributedCacheEntryOptions? options = null) where T : class
    {
        var data = await cache.GetObjectAsync<T>(key);
        if (data != null)
        {
            return data;
        }
        options ??= new DistributedCacheEntryOptions() { SlidingExpiration = TimeSpan.FromHours(24) };
        data = factory.Invoke(options);
        if (data != null)
        {
            await cache.SetObjectAsync(key, data, options);
        }
        return data;
    }
}
