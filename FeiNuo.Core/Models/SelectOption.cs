using System.ComponentModel;

namespace FeiNuo.Core;

/// <summary>
/// 前端下拉列表的选项模型
/// </summary>
public class SelectOption
{
    /// <summary>
    /// 选项值
    /// </summary>
    [Description("选项值")]
    public object Value { get; set; }

    /// <summary>
    /// 显示文本
    /// </summary>
    [Description("显示文本")]
    public string Label { get; set; }

    /// <summary>
    /// 是否禁用
    /// </summary>
    [Description("是否禁用")]
    public bool Disabled { get; set; }

    /// <summary>
    /// 颜色
    /// </summary>
    [Description("颜色")]
    public string? Color { get; set; } = null;

    /// <summary>
    /// 是否禁用
    /// </summary>
    [Description("是否可选择")]
    public bool? Selectable { get; set; } = null;

    /// <summary>
    /// 下级节点：树形结构使用
    /// </summary>
    [Description("下级节点：树形结构使用")]
    public List<SelectOption>? Children { get; set; }

    /// <summary>
    /// 其它内容：比如额外属性值，助词码等
    /// </summary>
    [Description("其它内容：比如额外属性值，助词码等")]
    public object? Data { get; set; }

    /// <summary>
    /// 构造函数
    /// </summary>
    public SelectOption(object value, string label, bool disabled = false)
    {
        Value = value;
        Label = label;
        Disabled = disabled;
    }
}