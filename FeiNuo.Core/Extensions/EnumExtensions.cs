using System.ComponentModel;
using System.Reflection;

namespace FeiNuo.Core;

/// <summary>
/// 枚举扩展
/// </summary>
public static class EnumExtensions
{
    /// <summary>
    /// 获取描述信息，枚举需要使用 [Description]属性
    /// </summary>
    /// <param name="value">枚举值</param>
    /// <returns>[Description]对应的内容，如果没有注解返回空字符串</returns>
    public static string GetDescription(this Enum value)
    {
        var field = value.GetType().GetField(value.ToString());
        if (field == null) return "";

        return field.GetCustomAttribute(typeof(DescriptionAttribute), false) is DescriptionAttribute attribute
            ? attribute.Description : "";
    }
}