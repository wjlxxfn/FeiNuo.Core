namespace FeiNuo.Core
{
    /// <summary>
    /// DTO 基类
    /// </summary>
    public class BaseDto : ICloneable
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
        /// 浅拷贝，调用Object.MemberwiseClone实现，如需深拷贝需另外实现
        /// </summary>
        public object Clone()
        {
            return MemberwiseClone();
        }
    }
}
