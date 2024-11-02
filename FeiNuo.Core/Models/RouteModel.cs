namespace FeiNuo.Core
{
    /// <summary>
    /// Vue路由模型
    /// </summary>
    [Serializable]
    public class RouteVO
    {
        /// <summary>
        /// 路由名
        /// </summary>
        public string Name { get; set; } = null!;

        /// <summary>
        /// 路径
        /// </summary>
        public string Path { get; set; } = null!;

        /// <summary>
        /// Vue组件
        /// </summary>
        public string? Component { get; set; }

        /// <summary>
        /// 重定向地址
        /// </summary>
        public string? Redirect { get; set; }

        /// <summary>
        /// 元数据
        /// </summary>
        public MetaVO Meta { get; set; } = null!;

        /// <summary>
        /// 子路由
        /// </summary>
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
        public string Title { get; set; } = null!;

        /// <summary>
        /// 图标
        /// </summary>
        public string? Icon { get; set; }

        /// <summary>
        /// 是否隐藏：不显示在菜单中
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// 禁止缓存页面
        /// </summary>
        public bool NoCache { get; set; }

        /// <summary>
        /// 默认构造函数
        /// </summary>
        public MetaVO() { }

        /// <summary>
        /// 构造函数
        /// </summary>
        public MetaVO(string title, string? icon, bool hidden = false, bool noCache = false)
        {
            Title = title;
            Hidden = hidden;
            Icon = icon;
            NoCache = noCache;
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
                return Path.StartsWith("http://", StringComparison.CurrentCultureIgnoreCase)
                    || Path.StartsWith("https://", StringComparison.CurrentCultureIgnoreCase);
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
}
