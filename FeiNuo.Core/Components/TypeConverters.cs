using System.ComponentModel;
using System.Globalization;

namespace FeiNuo.Core
{
    /// <summary>
    /// DateOnly，默认只支持日期格式，如果前端是完整的日期包含时间的转换不了，
    /// 这里先转成DateTime在转成DateOnly
    /// </summary>
    public class DateOnlyTypeConverter : DateOnlyConverter
    {
        public override object? ConvertFrom(ITypeDescriptorContext? context, CultureInfo? culture, object value)
        {
            if (value is string text && text.Trim().Length > 10)
            {
                try
                {
                    var dt = DateTime.Parse(text);
                    return DateOnly.FromDateTime(dt);
                }
                catch (FormatException e)
                {
                    throw new FormatException("日期转换错误", e);
                }
            }
            return base.ConvertFrom(context, culture, value);
        }
    }
}
