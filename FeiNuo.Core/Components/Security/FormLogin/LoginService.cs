using FeiNuo.Core.Utilities;
using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Configuration;

namespace FeiNuo.Core.Security
{
    public class LoginService : ILoginService
    {
        protected readonly ITokenService tokenService;
        protected readonly ILoginUserService userService;
        protected readonly IDistributedCache cache;
        protected readonly SecurityOptions securityOptions = new();

        public LoginService(ILoginUserService userService, ITokenService tokenService, IDistributedCache cache, IConfiguration configuration)
        {
            this.userService = userService;
            this.tokenService = tokenService;
            this.cache = cache;
            configuration.GetSection(SecurityOptions.ConfigKey).Bind(securityOptions);
        }

        public async Task<string> HandleLogin(LoginForm form)
        {
            // 验证码核对
            await CheckCaptcha(form);

            // 查询用户信息
            var user = await userService.LoadUserByUsername(form.Username);
            if (null == user || user.Username != form.Username) // 数据库不区分大小写的，这里在判断一次
            {
                throw new MessageException($"用户名【{form.Username}】不存在！");
            }
            if (!userService.ValidatePassword(form, user))
            {
                throw new MessageException($"用户名或密码错误！", MessageType.Error);
            }

            // 登录成功，清空使用的验证码缓存
            if (securityOptions.Captcha.Enabled && form.CaptchaKey != "")
            {
                cache.Remove(AppConstants.CACHE_PREFIX_CAPTCHA + form.CaptchaKey);
            }

            // 生成Token
            return await tokenService.CreateTokenAsync(user);
        }

        protected async Task CheckCaptcha(LoginForm form)
        {
            if (securityOptions.Captcha.Enabled)
            {
                if (string.IsNullOrEmpty(form.CaptchaKey)) throw new MessageException("缺少参数cptchaKey");
                var code = await cache.GetStringAsync(AppConstants.CACHE_PREFIX_CAPTCHA + form.CaptchaKey);
                if (string.IsNullOrEmpty(code))
                {
                    throw new MessageException("验证码不存在或已失效！", MessageType.Error);
                }
                if (!code.Equals(form.Captcha.ToLower(), StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new MessageException($"验证码错误！", MessageType.Error);
                }
            }
        }

        /// <summary>
        /// 退出登录：作废token,记录日志
        /// </summary>
        /// <param name="token">要退出的用户token</param>
        /// <param name="user">当前操作用户</param>
        public async Task HandleLogout(string token, LoginUser user)
        {
            // 作废token
            await tokenService.DisableTokenAsync(token);
        }

        /// <summary>
        /// 生成验证码
        /// </summary>
        public async Task<CaptchaResult> CreateCaptcha()
        {
            if (securityOptions.Captcha.Enabled)
            {
                var key = Guid.NewGuid().ToString();
                var captcha = CaptchaUtils.CreateCaptcha(securityOptions.Captcha);

                // 加入缓存
                var cacheOptions = new DistributedCacheEntryOptions
                {
                    AbsoluteExpirationRelativeToNow = TimeSpan.FromSeconds(securityOptions.Captcha.Timeout)
                };
                await cache.SetStringAsync(AppConstants.CACHE_PREFIX_CAPTCHA + key, captcha.Text, cacheOptions);
                return new CaptchaResult(key, captcha.ImageBase64);
            }
            return new CaptchaResult();
        }

        /// <summary>
        /// 获取登录用户信息
        /// </summary>
        public async Task<Dictionary<string, object>> GetLoginUserInfo(LoginUser user)
        {
            return await userService.GetLoginUserInfo(user);
        }
    }
}
