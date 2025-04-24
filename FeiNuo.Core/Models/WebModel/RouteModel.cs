using System.ComponentModel;

namespace FeiNuo.Core;

/// <summary>
/// Vue路由模型
/// </summary>
[Serializable]
public class RouteVO
{
    /// <summary>
    /// 路由名
    /// </summary>
    [Description("路由名")]
    public string Name { get; set; } = null!;

    /// <summary>
    /// 路径
    /// </summary>
    [Description("路径")]
    public string Path { get; set; } = null!;

    /// <summary>
    /// Vue组件
    /// </summary>
    [Description("Vue组件")]
    public string? Component { get; set; }

    /// <summary>
    /// 重定向地址
    /// </summary>
    [Description("重定向地址")]
    public string? Redirect { get; set; }

    /// <summary>
    /// 元数据
    /// </summary>
    [Description("元数据")]
    public MetaVO Meta { get; set; } = null!;

    /// <summary>
    /// 子路由
    /// </summary>
    [Description("子路由")]
    public List<RouteVO> Children { get; set; } = [];
}

/// <summary>
/// 路由元数据
/// </summary>
[Serializable]
public class MetaVO
{
    /// <summary>
    /// 标题
    /// </summary>
    [Description("标题")]
    public string Title { get; set; } = null!;

    /// <summary>
    /// 图标
    /// </summary>
    [Description("图标")]
    public string? Icon { get; set; }

    /// <summary>
    /// 是否隐藏：不显示在菜单中,默认false
    /// </summary>
    [Description("是否隐藏：不显示在菜单中,默认false")]
    public bool? Hidden { get; set; }

    /// <summary>
    /// 是否固定在页签上,默认false
    /// </summary>
    [Description("是否固定在页签上,默认false")]
    public bool? Affix { get; set; }

    /// <summary>
    /// 是否显示在面包屑导航中,默认true
    /// </summary>
    [Description("是否显示在面包屑导航中")]
    public bool? Breadcrumb { get; set; }

    /// <summary>
    /// 是否缓存页面,默认true
    /// </summary>
    [Description("是否缓存页面,默认true")]
    public bool? keepAlive { get; set; }

    /// <summary>
    /// 角色,多个用逗号分隔
    /// </summary>
    [Description("角色,多个用逗号分隔")]
    public string? Roles { get; set; }

    /// <summary>
    /// 权限,多个用逗号分隔
    /// </summary>
    [Description("权限,多个用逗号分隔")]
    public string? Permissions { get; set; }

    /// <summary>
    /// 默认构造函数
    /// </summary>
    public MetaVO() { }

    /// <summary>
    /// 构造函数
    /// </summary>
    public MetaVO(string title, string? icon, bool hidden = false)
    {
        Title = title;
        Icon = icon;
        Hidden = hidden;
    }
}

/// <summary>
/// 菜单结构
/// </summary>
public class MenuVO
{
    /// <summary>
    /// 菜单名称
    /// </summary>
    public string Title { get; set; } = null!;

    /// <summary>
    /// 菜单路径
    /// </summary>
    public string Path { get; set; } = null!;

    /// <summary>
    /// 菜单图标
    /// </summary>
    public string? Icon { get; set; }

    /// <summary>
    /// 是否外部链接
    /// </summary>
    public bool IsLink
    {
        get
        {
            return Path.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || Path.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// 下级菜单集合
    /// </summary>
    public IEnumerable<MenuVO> Children { get; set; } = [];

    /// <summary>
    /// 默认构造函数
    /// </summary>
    public MenuVO() { }

    /// <summary>
    /// 默认构造函数
    /// </summary>
    public MenuVO(string title, string path, string? icon = null)
    {
        Title = title;
        Path = path;
        Icon = icon;
    }
}
