namespace FeiNuo.Core
{
    /// <summary>
    /// 实体 基类
    /// </summary>
    public class BaseEntity
    {
        /// <summary>
        /// 创建人
        /// </summary>
        public string? CreatedBy { get; set; }

        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime? CreatedTime { get; set; }

        /// <summary>
        /// 修改人
        /// </summary>
        public string? UpdatedBy { get; set; }

        /// <summary>
        /// 修改时间
        /// </summary>
        public DateTime? UpdatedTime { get; set; }
    }
}
