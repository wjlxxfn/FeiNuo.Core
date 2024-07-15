namespace FeiNuo.Core;

/// <summary>
/// 系统常量
/// </summary>
public class AppConstants
{
    /// <summary>
    /// 超级管理员的特殊权限和角色
    /// </summary>
    public const string SUPER_ADMIN = "SuperAdmin";

    /// <summary>
    /// 开发环境超管固定token，不需要额外登录
    /// </summary>
    public const string SUPER_ADMIN_TOKEN = "00000000000000000000000000000000";

    #region 缓存键值前缀

    /// <summary>
    /// 会话token缓存前缀
    /// </summary>
    public const string CACHE_PREFIX_SESSION = "SESSION::";

    /// <summary>
    /// 验证码缓存前缀
    /// </summary>
    public const string CACHE_PREFIX_CAPTCHA = "CAPTCHA::";

    /// <summary>
    /// 字典数据缓存前缀
    /// </summary>
    public const string CACHE_PREFIX_DICT = "DICT::";

    /// <summary>
    /// 配置项缓存前缀
    /// </summary>
    public const string CACHE_PREFIX_CONFIG = "CONFIG::";

    /// <summary>
    /// TOKEN黑名单缓存名
    /// </summary>
    public const string CACHE_PREFIX_FORBIDDEN = "FORBIDDEN::";
    #endregion
}