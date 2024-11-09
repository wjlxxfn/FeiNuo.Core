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

        /// <summary>
        /// 取值逻辑
        /// </summary>
        public Func<object, object?> ValueGetter { get; set; } = o => o;

        /// <summary>
        /// 列的赋值方法:返回该列是否唯一列
        /// </summary>
        public Func<object, IConvertible?, string> ValueSetter { get; set; } = (o, v) => throw new NotImplementedException();

    }

    /// <summary>
    /// 指定数据类型的列配置,主要导出数据用，需设置valueGetter
    /// </summary>
    public class ExcelColumn<T> : ExcelColumn where T : class
    {
        public ExcelColumn(string title, int? width = null, string format = "") : base(title, width, format)
        {
        }

        public ExcelColumn(string title, Func<T, object?> valueGetter, int? width = null, string format = "") : base(title, width, format)
        {
            ValueGetter = o => valueGetter((T)o);
        }

        /// <summary>
        /// 配置样式：方便链式调用
        /// </summary>
        public new ExcelColumn<T> ConfigStyle(Action<ExcelStyle> styleConfig)
        {
            styleConfig.Invoke(ColumnStyle);
            return this;
        }
    }

    /// <summary>
    /// 指定数据类型的列配置:主要导入导出用，导入时必须设置valueSetter
    /// </summary>
    /// <typeparam name="T">数据的实体类型</typeparam>
    /// <typeparam name="D">Excel的数据类型</typeparam>
    public class ExcelColumn<T, D> : ExcelColumn<T> where T : class
    {
        public ExcelColumn(string title, int? width = null, string format = "") : base(title, width, format)
        {
        }

        public ExcelColumn(string title, Func<T, object?> valueGetter, int? width = null, string format = "") : base(title, valueGetter, width, format)
        {
            ValueGetter = o => valueGetter((T)o);
        }

        public ExcelColumn(string title, Action<T, D?> valueSetter, bool required = false, int? width = null, string format = "") : this(title, width, format)
        {
            Required = required;
            ValueSetter = (o, v) =>
            {
                try
                {
                    // 数据转换
                    var val = (D?)Convert.ChangeType(v, typeof(D));
                    if ((val == null || string.IsNullOrWhiteSpace(val.ToString())))
                    {
                        if (required) return "不能为空";
                        else // 空值直接调用赋值方法，由调用法处理 比如 u.Username = v ?? "";
                        {
                            valueSetter.Invoke((T)o, val);
                            // 如果是唯一键值，返回_UniqueKey_特殊处理
                            return UniqueKey ? "_UniqueKey_" : "";
                        }
                    }

                    // 数据效验
                    var msg = InternalValidator?.Invoke(val) ?? "";
                    if (!string.IsNullOrWhiteSpace(msg)) return msg;

                    if (LimitValues.Any() && !LimitValues.Contains(val))
                    {
                        return $"可选值:{string.Join(",", LimitValues)}";
                    }

                    msg = Validator?.Invoke(val) ?? "";
                    if (!string.IsNullOrWhiteSpace(msg)) return msg;

                    // 赋值操作
                    valueSetter((T)o, val);

                    // 如果是唯一键值，返回_UniqueKey_特殊处理
                    return UniqueKey ? "_UniqueKey_" : "";
                }
                catch (Exception ex)
                {
                    return "数据类型转换错误：" + ex.Message;
                }
            };
        }

        /// <summary>
        /// 导入用：是否必填
        /// </summary>
        public bool Required { get; set; } = false;

        /// <summary>
        /// 导入用：是否唯一键，系统会把所有唯一键字段用竖线(|)拼接起来做为唯一标识，判断excel中是否有重复的值，如果有提示重复
        /// </summary>
        public bool UniqueKey { get; set; } = false;

        /// <summary>
        /// 导入用：验证值的合法性
        /// </summary>
        public Func<D, string>? Validator { get; set; }

        /// <summary>
        /// 导入用：可选值
        /// </summary>
        public IEnumerable<D> LimitValues { get; set; } = [];

        /// <summary>
        /// 导入用：子类配置的内容效验规则
        /// </summary>
        protected Func<D, string>? InternalValidator { get; set; }

        /// <summary>
        /// 配置样式：方便链式调用
        /// </summary>
        public new ExcelColumn<T, D> ConfigStyle(Action<ExcelStyle> styleConfig)
        {
            styleConfig.Invoke(ColumnStyle);
            return this;
        }
    }



    public class ExcelColumnString<T> : ExcelColumn<T, string> where T : class, new()
    {
        private const string FORMAT = "@";
        public ExcelColumnString(string title, int? width = null) : base(title, width, FORMAT)
        {
        }

        public ExcelColumnString(string title, Action<T, string?> valueSetter, bool required = false, int? width = null) : base(title, valueSetter, required, width, FORMAT)
        {
        }

        public ExcelColumnString(string title, Func<T, object?> valueGetter, int? width = null) : base(title, valueGetter, width, FORMAT)
        {
        }
    }

    public class ExcelColumnDate<T> : ExcelColumn<T, DateOnly?> where T : class, new()
    {
        private const int WIDTH = 13;
        private const string FORMAT = "yyyy-mm-dd";
        public ExcelColumnDate(string title, int? width = WIDTH) : base(title, width, FORMAT)
        {
        }

        public ExcelColumnDate(string title, Func<T, object?> valueGetter, int? width = WIDTH) : base(title, valueGetter, width)
        {
            ColumnStyle.DataFormat = FORMAT;
            ColumnStyle.HorizontalAlignment = 2;
        }

        public ExcelColumnDate(string title, Action<T, DateOnly?> valueSetter, bool required = false, int? width = WIDTH) : base(title, valueSetter, required, width)
        {
            ColumnStyle.DataFormat = FORMAT;
            ColumnStyle.HorizontalAlignment = 2;
        }
    }
    public class ExcelColumnDateTime<T> : ExcelColumn<T, DateTime?> where T : class, new()
    {
        private const int WIDTH = 17;
        private const string FORMAT = "yyyy-mm-dd hh:mm";

        public ExcelColumnDateTime(string title, int? width = WIDTH) : base(title, width, FORMAT)
        {
        }

        public ExcelColumnDateTime(string title, Func<T, object?> valueGetter, int? width = WIDTH) : base(title, valueGetter, width)
        {
            ColumnStyle.DataFormat = FORMAT;
            ColumnStyle.HorizontalAlignment = 2;
        }

        public ExcelColumnDateTime(string title, Action<T, DateTime?> valueSetter, bool required = false, int? width = WIDTH) : base(title, valueSetter, required, width)
        {
            ColumnStyle.DataFormat = FORMAT;
            ColumnStyle.HorizontalAlignment = 2;
        }
    }
    public class ExcelColumnTime<T> : ExcelColumn<T, DateTime?> where T : class, new()
    {
        private const int WIDTH = 10;
        private const string FORMAT = "hh:mm";

        public ExcelColumnTime(string title, int? width = WIDTH) : base(title, width, FORMAT)
        {
        }

        public ExcelColumnTime(string title, Func<T, object?> valueGetter, int? width = WIDTH) : base(title, valueGetter, width)
        {
            ColumnStyle.DataFormat = FORMAT;
            ColumnStyle.HorizontalAlignment = 2;
        }

        public ExcelColumnTime(string title, Action<T, DateTime?> valueSetter, bool required = false, int? width = WIDTH) : base(title, valueSetter, required, width)
        {
            ColumnStyle.DataFormat = FORMAT;
            ColumnStyle.HorizontalAlignment = 2;
        }
    }

    public class ExcelColumnInteger<T> : ExcelColumn<T, int?> where T : class, new()
    {
        private const int WIDTH = 8;

        public ExcelColumnInteger(string title, int? width = WIDTH) : base(title, width)
        {
        }

        public ExcelColumnInteger(string title, Func<T, object?> valueGetter, int? width = WIDTH) : base(title, valueGetter, width)
        {
        }

        public ExcelColumnInteger(string title, Action<T, int?> valueSetter, bool required = false, int? width = WIDTH) : base(title, valueSetter, required, width)
        {
            InternalValidator = v =>
            {
                if (MinValue.HasValue && v < MinValue) return $"必须大于等于：{MinValue.Value:0.######}";
                if (MaxValue.HasValue && v > MaxValue) return $"必须小于等于：{MaxValue.Value:0.######}";
                return "";
            };
        }

        public int? MinValue { get; set; }
        public int? MaxValue { get; set; }
    }

    public class ExcelColumnDecimal<T> : ExcelColumn<T, decimal?> where T : class, new()
    {
        private const int WIDTH = 8;

        public ExcelColumnDecimal(string title, int? width = WIDTH) : base(title, width)
        {
        }

        public ExcelColumnDecimal(string title, Func<T, object?> valueGetter, int? width = WIDTH) : base(title, valueGetter, width)
        {
        }

        public ExcelColumnDecimal(string title, Action<T, decimal?> valueSetter, bool required = false, int? width = WIDTH) : base(title, valueSetter, required, width)
        {
            InternalValidator = v =>
            {
                if (MinValue.HasValue && v < MinValue) return $"必须大于等于：{MinValue.Value:0.######}";
                if (MaxValue.HasValue && v > MaxValue) return $"必须小于等于：{MaxValue.Value:0.######}";
                return "";
            };
        }

        public decimal? MinValue { get; set; }
        public decimal? MaxValue { get; set; }
    }

    public class ExcelColumnPersent<T> : ExcelColumnDecimal<T> where T : class, new()
    {
        private const int WIDTH = 8;
        private const string FORMAT = "0.00%";

        public ExcelColumnPersent(string title, int? width = WIDTH) : base(title, width)
        {
            ColumnStyle.DataFormat = FORMAT;
        }

        public ExcelColumnPersent(string title, Func<T, object?> valueGetter, int? width = WIDTH) : base(title, valueGetter, width)
        {
            ColumnStyle.DataFormat = FORMAT;
        }

        public ExcelColumnPersent(string title, Action<T, decimal?> valueSetter, bool required = false, int? width = 8) : base(title, valueSetter, required, width)
        {
            ColumnStyle.DataFormat = FORMAT;
            MinValue = 0; MaxValue = 1;

            InternalValidator = v =>
            {
                if (MinValue.HasValue && v < MinValue) return $"必须大于等于：{MinValue.Value:0.######}";
                if (MaxValue.HasValue && v > MaxValue) return $"必须小于等于：{MaxValue.Value:0.######}";
                return "";
            };
        }
    }
}
