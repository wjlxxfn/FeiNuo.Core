namespace FeiNuo.Core;

/// <summary>
/// 通用更新模型：适用于更新少量常见字段，方便传参用
/// </summary>
public class UpdateDto : BaseDto
{
    /// <summary>
    /// 业务ID
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// 更新状态
    /// </summary>
    public StatusEnum? Status { get; set; }

    /// <summary>
    /// 更新状态
    /// </summary>
    public int? State { get; set; }

    /// <summary>
    /// 更新内容
    /// </summary>
    public string? Data { get; set; }

    /// <summary>
    /// 更新内容
    /// </summary>
    public string? Content { get; set; }

    /// <summary>
    /// 其他内容
    /// </summary>
    public string? ExtData { get; set; }

    /// <summary>
    /// 更新数量
    /// </summary>
    public decimal? Quantity { get; set; }

    /// <summary>
    /// 更新金额
    /// </summary>
    public decimal? Amount { get; set; }

    /// <summary>
    /// 更新类型
    /// </summary>
    public int? Type { get; set; }

    /// <summary>
    /// 更新数量
    /// </summary>
    public int? Count { get; set; }

    /// <summary>
    /// 更新类型
    /// </summary>
    public string? ActionType { get; set; }

    /// <summary>
    /// 更新值
    /// </summary>
    public string? ActionValue { get; set; }

    /// <summary>
    /// 更新日期
    /// </summary>
    public DateTime? Date { get; set; }

}
