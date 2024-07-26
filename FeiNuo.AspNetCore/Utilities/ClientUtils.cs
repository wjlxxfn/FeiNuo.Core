using Microsoft.AspNetCore.Http;

namespace FeiNuo.Core.Utilities
{
    public class ClientUtils
    {
        /// <summary>
        /// 获取客户端信息，使用UAParser
        /// </summary>
        public static RequestClient GetClientInfo(HttpContext context)
        {
            var ua = UAParser.Parser.GetDefault().Parse(context.Request.Headers.UserAgent.ToString());

            var os = ua.OS.ToString();
            var browser = ua.UA.ToString();
            var device = ua.Device.ToString();
            var isMobile = Mobile_OS.Contains(ua.OS.Family)
                || Mobile_Browser.Contains(ua.UA.Family)
                || Mobile_Device.Contains(ua.Device.Family);

            var ip = GetIpAddress(context);
            return new RequestClient(ip, os, browser, device, isMobile);
        }

        /// <summary>
        /// 获取IP地址
        /// Program.cs配置，且该中间件要加到最前面
        /// app.UseForwardedHeaders(new ForwardedHeadersOptions { ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto });
        /// nginx配置
        /// proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        /// proxy_set_header X-Forwarded-Proto $scheme;
        /// </summary>
        public static string GetIpAddress(HttpContext context)
        {
            var ipAddress = context.Request.Headers["X-Forwarded-For"].FirstOrDefault();
            if (string.IsNullOrEmpty(ipAddress))
            {
                var ip = context.Connection.RemoteIpAddress;
                ipAddress = ip == null ? "" : ip.MapToIPv4().ToString();
            }
            return ipAddress;

            // Program.cs配置，且该中间件要加到最前面
            // app.UseForwardedHeaders(new ForwardedHeadersOptions { ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto });

            // nginx配置
            // proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
            // proxy_set_header X-Forwarded-Proto $scheme;
        }

        #region Mobile UAs, OS & Devices
        private static readonly HashSet<string> Mobile_OS =
        [
            "Android",
            "iOS",
            "Windows Mobile",
            "Windows Phone",
            "Windows CE",
            "Symbian OS",
            "BlackBerry OS",
            "BlackBerry Tablet OS",
            "Firefox OS",
            "Brew MP",
            "webOS",
            "Bada",
            "Kindle",
            "Maemo"
        ];

        private static readonly HashSet<string> Mobile_Browser =
        [
            "Android",
            "Firefox Mobile",
            "Opera Mobile",
            "Opera Mini",
            "Mobile Safari",
            "Amazon Silk",
            "webOS Browser",
            "MicroB",
            "Ovi Browser",
            "NetFront",
            "NetFront NX",
            "Chrome Mobile",
            "Chrome Mobile iOS",
            "UC Browser",
            "Tizen Browser",
            "Baidu Explorer",
            "QQ Browser Mini",
            "QQ Browser Mobile",
            "IE Mobile",
            "Polaris",
            "ONE Browser",
            "iBrowser Mini",
            "Nokia Services (WAP) Browser",
            "Nokia Browser",
            "Nokia OSS Browser",
            "BlackBerry WebKit",
            "BlackBerry",
            "Palm",
            "Palm Blazer",
            "Palm Pre",
            "Teleca Browser",
            "SEMC-Browser",
            "PlayStation Portable",
            "Nokia",
            "Maemo Browser",
            "Obigo",
            "Bolt",
            "Iris",
            "UP.Browser",
            "Minimo",
            "Bunjaloo",
            "Jasmine",
            "Dolfin",
            "Polaris",
            "Skyfire"
        ];

        private static readonly HashSet<string> Mobile_Device =
        [
            "BlackBerry",
            "MI PAD",
            "iPhone",
            "iPad",
            "iPod",
            "Kindle",
            "Kindle Fire",
            "Nokia",
            "Lumia",
            "Palm",
            "DoCoMo",
            "HP TouchPad",
            "Xoom",
            "Motorola",
            "Generic Feature Phone",
            "Generic Smartphone"
        ];
        #endregion
    }
}
