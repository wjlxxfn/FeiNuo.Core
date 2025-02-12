using SixLabors.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Drawing.Processing;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;

namespace FeiNuo.Core.Captcha;

internal class CaptchaGenerator
{
    private readonly CaptchaOptions options;
    private readonly Color[] Colors = [Color.Black, Color.Red, Color.Blue, Color.Green, Color.Orange, Color.Brown, Color.Cyan, Color.Gray, Color.DarkRed, Color.DarkOrange, Color.DarkBlue];
    public CaptchaGenerator(CaptchaOptions options)
    {
        this.options = options;
    }

    /// <summary>
    /// 生成验证码图片
    /// </summary>
    /// <param name="captchaText">验证码内容</param>
    public CaptchaData GetCaptcha(string captchaText)
    {
        if (string.IsNullOrEmpty(captchaText)) throw new ArgumentException("必须指定验证码内容");

        var rand = new Random();

        FontFamily fontFamily;
        #region 获取字体,linux下可能没有字体
        if (SystemFonts.Families.Any()) // 存在系统字体，随机取一种
        {
            var index = rand.Next(SystemFonts.Families.Count() - 1);
            fontFamily = SystemFonts.Families.ElementAt(index);
        }
        else  //linux下没有对应的字体文件，把字体文件放到项目里
        {
            var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fonts");
            var dirFonts = Directory.GetFiles(dir, "*.ttf", SearchOption.AllDirectories);
            if (dirFonts.Length == 0)
            {
                throw new MessageException("找不到字体文件。");
            }
            var index = rand.Next(dirFonts.Length - 1);
            fontFamily = new FontCollection().Add(dirFonts[index]);
        }
        #endregion

        var captcha = new CaptchaData() { Text = captchaText, ImageFormat = options.ImageType.ToString().ToLower() };
        using (var image = new Image<Rgba32>(options.Width, options.Height, Color.Parse(options.BackGroundColor)))
        {
            // 字体
            image.Mutate(ctx =>
            {
                // 画验证码
                for (int i = 0; i < captchaText.Length; i++)
                {
                    float y = rand.Next(1, 8);
                    float x = 1f * (options.Width - 10) / captchaText.Length * i + 5;

                    var fontSize = rand.Next(Math.Max(25, options.FontSize - 10), Math.Min(options.FontSize + 10, 45));
                    var font = fontFamily.CreateFont(fontSize, (FontStyle)rand.Next(4));
                    var color = Colors[rand.Next(Colors.Length)];
                    ctx.DrawText(captchaText[i].ToString(), font, color, new PointF(x, y));
                }

                // 画干扰线
                if (options.LineCount > 0)
                {
                    for (int i = 0; i < options.LineCount; i++)
                    {
                        var pen = Pens.Solid(Colors[rand.Next(Colors.Length)], 1);
                        var p1 = new PointF(rand.Next(options.Width), rand.Next(options.Height));
                        var p2 = new PointF(rand.Next(options.Width), rand.Next(options.Height));
                        ctx.DrawLine(pen, p1, p2);
                    }
                }

                // 画噪点
                if (options.ChaosDensity > 0)
                {
                    for (int i = 0; i < options.Width * options.Height / options.ChaosDensity; i++)
                    {
                        var pen = Pens.Solid(Colors[rand.Next(Colors.Length)], 1);
                        var p1 = new PointF(rand.Next(options.Width), rand.Next(options.Height));
                        var p2 = new PointF(p1.X + 1f, p1.Y + 1f);
                        ctx.DrawLine(pen, p1, p2);
                    }
                }
            });
            using var ms = new MemoryStream();
            switch (options.ImageType.ToLower())
            {
                case "png":
                    image.SaveAsPng(ms);
                    captcha.ImageFormat = "png";
                    break;
                case "jpg":
                case "jpeg":
                    captcha.ImageFormat = "jpeg";
                    image.SaveAsJpeg(ms);
                    break;
                case "bmp":
                    captcha.ImageFormat = "bmp";
                    image.SaveAsBmp(ms);
                    break;
                default:
                    captcha.ImageFormat = "gif";
                    image.SaveAsGif(ms);
                    break;
            }
            captcha.ImageBytes = ms.ToArray();
        }
        return captcha;
    }
}
