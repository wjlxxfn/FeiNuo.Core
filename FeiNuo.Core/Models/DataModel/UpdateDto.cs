namespace FeiNuo.Core;

/// <summary>
/// 通用更新模型：适用于更新少量常见字段，方便传参用
/// </summary>
public class UpdateDto : BaseDto
{
    /// <summary>
    /// Id
    /// </summary>
    public int? Id { get; set; }

    /// <summary>
    /// Id
    /// </summary>
    public string? Sid { get; set; }

    /// <summary>
    /// 名称
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// 编号
    /// </summary>
    public string? BillNo { get; set; }

    /// <summary>
    /// 键
    /// </summary>
    public string? Key { get; set; }

    /// <summary>
    /// 值
    /// </summary>
    public string? Value { get; set; }

    /// <summary>
    /// 标签
    /// </summary>
    public string? Label { get; set; }

    /// <summary>
    /// 状态
    /// </summary>
    public int? Status { get; set; }

    /// <summary>
    /// 状态
    /// </summary>
    public int? State { get; set; }

    /// <summary>
    ///  是否启用
    /// </summary>
    public bool? Enabled { get; set; }

    /// <summary>
    /// 是否禁用
    /// </summary>
    public bool? Disabled { get; set; }

    /// <summary>
    /// 数据
    /// </summary>
    public string? Data { get; set; }

    /// <summary>
    /// 内容
    /// </summary>
    public string? Content { get; set; }

    /// <summary>
    /// 其他内容
    /// </summary>
    public string? ExtData { get; set; }

    /// <summary>
    /// 数量
    /// </summary>
    public decimal? Quantity { get; set; }

    /// <summary>
    /// 税额
    /// </summary>
    public decimal? Tax { get; set; }

    /// <summary>
    /// 单价
    /// </summary>
    public decimal? Price { get; set; }

    /// <summary>
    /// 金额
    /// </summary>
    public decimal? Amount { get; set; }

    /// <summary>
    /// 更新类型
    /// </summary>
    public int? Type { get; set; }

    /// <summary>
    /// 数量
    /// </summary>
    public int? Count { get; set; }

    /// <summary>
    /// 操作类型
    /// </summary>
    public string? ActionType { get; set; }

    /// <summary>
    /// 操作值
    /// </summary>
    public string? ActionValue { get; set; }

    /// <summary>
    /// 备注
    /// </summary>
    public string? Remark { get; set; }

    /// <summary>
    /// 描述
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// 日期
    /// </summary>
    public DateTime? Date { get; set; }

    /// <summary>
    /// 操作日期
    /// </summary>
    public DateTime? ActionTime { get; set; }

}
public class SimpleDto : UpdateDto { }
public class CommonDto : UpdateDto { }
