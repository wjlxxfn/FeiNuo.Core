using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Configuration;

namespace FeiNuo.AspNetCore.Security.Authentication;

public class CacheTokenService : ITokenService
{
    private readonly IDistributedCache cache;
    private readonly SecurityOptions cfg = new();
    private readonly AppConfig appCfg = new();
    public CacheTokenService(IDistributedCache cache, IConfiguration configuration)
    {
        this.cache = cache;
        configuration.GetSection(SecurityOptions.ConfigKey).Bind(cfg);
        configuration.GetSection(AppConfig.ConfigKey).Bind(appCfg);
    }

    public async Task<string> CreateTokenAsync(LoginUser user)
    {
        // 生成token
        var token = CacheTokenPrefix + Guid.NewGuid().ToString().Replace("-", "");

        // 保存到缓存，设置滑动更新，连续访问不需要刷新token
        var cacheOptions = new DistributedCacheEntryOptions
        {
            SlidingExpiration = TimeSpan.FromSeconds(cfg.TokenExpiration)
        };
        await cache.SetObjectAsync(token, user, cacheOptions);

        // 返回token
        return token;
    }

    public async Task<TokenValidationResult> ValidateTokenAsync(string token)
    {
        if (string.IsNullOrEmpty(token))
        {
            throw new ArgumentNullException(nameof(token));
        }

        if (!token.StartsWith(CacheTokenPrefix) || token.Length != CacheTokenPrefix.Length + 32)
        {
            return new TokenValidationResult("Token格式错误。");
        }

        // 从缓存中取数据
        var user = await cache.GetObjectAsync<LoginUser>(token);

        return user != null
            ? new TokenValidationResult(user)
            : new TokenValidationResult("Token不存在或已失效。");
    }

    public async Task DisableTokenAsync(string token)
    {
        await cache.RemoveAsync(token);
    }

    /// <summary>
    /// 缓存键值前缀：SESSION::AppCode:
    /// </summary>
    private string CacheTokenPrefix
    {
        get
        {
            return AppConstants.CACHE_PREFIX_SESSION + appCfg.AppCode + ":";
        }
    }
}
