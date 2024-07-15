namespace FeiNuo.Core
{
    /// <summary>
    /// DTO 基类
    /// </summary>
    public class BaseDto
    {
        /// <summary>
        /// 创建人
        /// </summary>
        public string? CreatedBy { get; set; }

        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime? CreatedTime { get; set; }
    }
}
