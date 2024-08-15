namespace FeiNuo.Core
{
    /// <summary>
    /// 下载excel的辅助类
    /// </summary>
    public class ExcelDownload
    {
        public string FileName { get; set; }
        public string ContentType { get; set; }
        public byte[] Bytes { get; set; }

        public ExcelDownload(string fileName, string contentType, byte[] bytes)
        {
            FileName = fileName;
            ContentType = contentType;
            Bytes = bytes;
        }
    }
}
