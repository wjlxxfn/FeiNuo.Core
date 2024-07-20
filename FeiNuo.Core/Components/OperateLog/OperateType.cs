using System.ComponentModel;

namespace FeiNuo.Core
{
    /// <summary>
    /// 操作类型枚举
    /// </summary>
    public enum OperateType
    {
        /// <summary>
        /// 未知类型
        /// </summary>
        [Description("未知类型")]
        Default = 0,

        /// <summary>
        /// 用户登录
        /// </summary>
        [Description("用户登录")]
        Login = 1,

        /// <summary>
        /// 退出登录
        /// </summary>
        [Description("退出登录")]
        Logout = 2,

        /// <summary>
        /// 新增操作
        /// </summary>
        [Description("新增操作")]
        Create = 3,

        /// <summary>
        /// 更新操作
        /// </summary>
        [Description("更新操作")]
        Update = 4,

        /// <summary>
        /// 删除操作
        /// </summary>
        [Description("删除操作")]
        Delete = 5,

        /// <summary>
        /// 查询操作
        /// </summary>
        [Description("查询操作")]
        Read = 6,

        /// <summary>
        /// 安全认证
        /// </summary>
        [Description("安全认证")]
        Security = 7,

        /// <summary>
        /// 授权操作
        /// </summary>
        [Description("授权操作")]
        Grant = 8,

        /// <summary>
        /// 其它操作
        /// </summary>
        [Description("其它操作")]
        Revoke = 10,
    }
}
