namespace FeiNuo.Core
{
    /// <summary>
    /// 前端下拉列表的选项模型
    /// </summary>
    public class SelectOption
    {
        /// <summary>
        /// 选项值
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// 显示文本
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 是否禁用
        /// </summary>
        public bool Disabled { get; set; }

        /// <summary>
        /// 是否禁用
        /// </summary>
        public bool? Selectable { get; set; } = null;

        /// <summary>
        /// 下级节点：树形结构使用
        /// </summary>
        public List<SelectOption>? Children { get; set; }

        /// <summary>
        /// 其它内容：比如额外属性值，助词码等
        /// </summary>
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
}