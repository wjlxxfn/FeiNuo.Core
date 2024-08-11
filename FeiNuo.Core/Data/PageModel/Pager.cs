using System.Text.Json.Serialization;

namespace FeiNuo.Core
{
    #region 分页对象 Pager
    /// <summary>
    /// 分页对象：需要页索引，每页记录数，以及排序信息
    /// </summary>
    public class Pager
    {
        /// <summary>
        /// 当前页码，从0开始
        /// </summary>
        public int PageNo { get; set; }

        /// <summary>
        /// 每页记录数
        /// </summary>
        public int PageSize { get; set; }

        /// <summary>
        /// 排序
        /// </summary>
        public List<SortItem> SortBy { get; set; } = [];

        /// <summary>
        /// 构造函数
        /// </summary>
        public Pager() { }

        /// <summary>
        /// 构造函数
        /// </summary>
        public Pager(int pageNo, int pageSize)
        {
            PageNo = pageNo;
            PageSize = pageSize;
        }

        /// <summary>
        /// 返回一个不分页的分页对象
        /// </summary>
        [JsonIgnore]
        public static Pager Unpaged => new() { PageSize = 0 };
    }
    #endregion


    /// <summary>
    /// 排序对象
    /// </summary>
    public class SortItem
    {
        /// <summary>
        /// 排序字段
        /// </summary>
        public string SortField { get; set; } = string.Empty;

        /// <summary>
        /// 排序类型
        /// </summary>
        public SortTypeEnum SortType { get; set; } = SortTypeEnum.ASC;

        /// <summary>
        /// 无参构造函数
        /// </summary>
        public SortItem()
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        public SortItem(string sortField, SortTypeEnum sortType = SortTypeEnum.ASC)
        {
            SortField = sortField;
            SortType = sortType;
        }
    }

    /// <summary>
    /// 排序类型枚举
    /// </summary>
    public enum SortTypeEnum
    {
        /// <summary>
        /// 升序
        /// </summary>
        ASC,

        /// <summary>
        /// 降序
        /// </summary>
        DESC
    }
}