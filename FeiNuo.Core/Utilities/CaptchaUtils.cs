using FeiNuo.Core.Captcha;

namespace FeiNuo.Core.Utilities
{
    public class CaptchaUtils
    {
        /// <summary>
        /// 生成验证码
        /// </summary>
        /// <param name="codeLen">字符长度</param>
        /// <param name="width">图片宽度</param>
        /// <param name="height">图片高度</param>
        /// <returns>Captcha</returns>
        public static CaptchaData CreateCaptcha(int codeLen = 4, int width = 100, int height = 40)
        {
            var options = new CaptchaOptions
            {
                Length = codeLen,
                Width = width,
                Height = height
            };
            return CreateCaptcha(options);
        }


        /// <summary>
        /// 生成验证码
        /// </summary>
        public static CaptchaData CreateCaptcha(CaptchaOptions options)
        {
            var text = CreateVerifyCode(options.Length);
            var gen = new CaptchaGenerator(options);
            return gen.GetCaptcha(text);
        }

        /// <summary>
        /// 生成随机验证码
        /// </summary>
        /// <param name="codeLength">验证码长度，默认4位</param>
        /// <param name="codeSerials">生成验证码的可选数据</param>
        public static string CreateVerifyCode(int codeLength = 4, string codeSerials = "2,3,4,5,6,7,8,9,a,b,c,e,f,g,h,k,m,n,p,q,r,s,t,u,v,w,x,y,z,A,B,C,E,F,G,H,K,M,N,P,Q,R,S,T,U,V,W,X,Y,Z")
        {
            var codes = codeSerials.Split(',');
            var rand = new Random();
            string code = "";
            for (int i = 0; i < codeLength; i++)
            {
                code += codes[rand.Next(codes.Length)];
            }
            return code;
        }
    }
}
