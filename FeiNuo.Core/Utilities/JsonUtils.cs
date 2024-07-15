using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;

namespace FeiNuo.Core.Utilities
{
    /// <summary>
    /// JSON 转换工具类
    /// </summary>
    public class JsonUtils
    {
        /// <summary>
        /// 序列化
        /// </summary>
        public static string Serialize(object obj, bool writeIndented = false)
        {
            var options = new JsonSerializerOptions();
            MergeSerializerOptions(options);
            options.WriteIndented = writeIndented;
            return JsonSerializer.Serialize(obj, options);
        }

        /// <summary>
        /// 反序列化
        /// </summary>
        public static T? Deserialize<T>(string json)
        {
            var options = new JsonSerializerOptions();
            MergeSerializerOptions(options);
            return JsonSerializer.Deserialize<T>(json, options);
        }

        /// <summary>
        /// 配置序列化选项
        /// </summary>
        /// <param name="options"></param>
        public static void MergeSerializerOptions(JsonSerializerOptions options)
        {
            // 空值不输出
            options.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull;
            // 阻断循环引用
            options.ReferenceHandler = ReferenceHandler.IgnoreCycles;
            // 设置支持的编码,防止转义
            options.Encoder = JavaScriptEncoder.Create(UnicodeRanges.All);
            // 小驼峰转换
            options.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
            // 支持从字符转成数字
            options.NumberHandling = JsonNumberHandling.AllowReadingFromString;

            // JSON格式化配置，日期格式 yyyy-MM-dd HH:mm:ss
            options.Converters.Add(new DateTimeConverter());
        }
    }

    #region JSON转换器
    /// <summary>
    /// 日期类型JSON处理,统一格式 yyyy-MM-dd HH:mm:ss
    /// </summary>
    internal class DateTimeConverter : JsonConverter<DateTime>
    {
        /// <summary>
        /// 转成DateTime
        /// </summary>
        public override DateTime Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (!reader.TryGetDateTime(out DateTime value))
            {
                value = DateTime.Parse(reader.GetString()!);
            }
            return value;
        }

        /// <summary>
        /// 转成yyyy-MM-dd HH:mm:ss格式
        /// </summary>
        public override void Write(Utf8JsonWriter writer, DateTime value, JsonSerializerOptions options)
        {
            if (value == DateTime.MinValue || value == DateTime.Parse("1900-01-01"))
            {
                writer.WriteNullValue();
            }
            else writer.WriteStringValue(value.ToString("yyyy-MM-dd HH:mm:ss"));
        }
    }
    #endregion

}
