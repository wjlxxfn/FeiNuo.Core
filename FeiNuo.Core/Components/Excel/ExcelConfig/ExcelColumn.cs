namespace FeiNuo.Core
{
    public class ExcelColumn
    {
        public ExcelColumn(string title, int? width = null, string format = "")
        {
            Title = title;
            Width = width;
            if (!string.IsNullOrWhiteSpace(format))
            {
                ColumnStyle.DataFormat = format;
            }
        }

        /// <summary>
        /// 标题名称：不允许重复， 如果多行，可根据#号分隔不同行标题<br/>
        /// 分行的情况下 列标题要么不带#号系统默认多行合并，要么#号数量必须一致，<br/>
        /// 系统先根据列标题是否一致合并多行，然后根据行标题是否相同来是否合列<br/>
        /// 比如列配置【姓名,角色#编码,角色#名称】：系统将生成两行标题，第一列姓名对应的两行合并， 第一行角色对应的两列合并。
        /// </summary>
        public string Title { get; set; } = null!;

        /// <summary>
        /// 标题拆分成多行标题
        /// </summary>
        public string[] RowTitles { get { return Title.Trim().Split('#'); } }

        /// <summary>
        /// 列宽度
        /// </summary>
        public int? Width { get; set; }

        /// <summary>
        /// 该列是否隐藏
        /// </summary>
        public bool Hidden { get; set; } = false;

        /// <summary>
        /// 列的样式
        /// </summary>
        public ExcelStyle ColumnStyle { get; } = new();

        /// <summary>
        /// 配置样式：方便链式调用
        /// </summary>
        public ExcelColumn ConfigStyle(Action<ExcelStyle> styleConfig)
        {
            styleConfig.Invoke(ColumnStyle);
            return this;
        }

    }

    /// <summary>
    /// 指定数据类型的列配置
    /// </summary>
    public class ExcelColumn<T> : ExcelColumn where T : class
    {
        public ExcelColumn(string title, Func<T, object?>? valueGetter, int? width = null, string format = "") : base(title, width, format)
        {
            ValueGetter = valueGetter;
        }
        public ExcelColumn(string title, Action<T, IConvertible?>? valueSetter, int? width = null, string format = "") : base(title, width, format)
        {
            ValueSetter = valueSetter;
        }

        /// <summary>
        /// 列的取值方法
        /// </summary>
        public Func<T, object?>? ValueGetter { get; set; }


        /// <summary>
        /// 列的赋值方法
        /// </summary>
        public Action<T, IConvertible?>? ValueSetter { get; set; }

        /// <summary>
        /// 配置样式：方便链式调用
        /// </summary>
        public new ExcelColumn<T> ConfigStyle(Action<ExcelStyle> styleConfig)
        {
            styleConfig.Invoke(ColumnStyle);
            return this;
        }
    }
}
