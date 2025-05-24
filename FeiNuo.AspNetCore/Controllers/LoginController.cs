using FeiNuo.Core.Captcha;
using FeiNuo.Core.Login;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace FeiNuo.AspNetCore.Controllers;

[Authorize]
[Route("/")]
public class LoginController : BaseController
{
    private readonly ILogService logService;
    private readonly ILoginService loginService;
    public LoginController(ILoginService loginService, ILogService logService)
    {
        this.loginService = loginService;
        this.logService = logService;
    }

    /// <summary>
    /// 登录系统
    /// </summary>
    /// <returns>认证token</returns>
    [AllowAnonymous]
    [SaveLog("用户登录", OperateType.Login, true, true)]
    [HttpPost("login")]
    public async Task<string> HandleLogin([FromBody] LoginForm form)
    {
        return await loginService.HandleLogin(form);
    }

    /// <summary>
    /// 退出登录
    /// </summary>
    [AllowAnonymous]
    [SaveLog("退出登录", OperateType.Logout, true, true)]
    [HttpPost("logout")]
    public async void HandleLogout()
    {
        var token = await GetAccessToken();
        if (string.IsNullOrWhiteSpace(token) || User == null)
        {
            return;
        }
        await loginService.HandleLogout(token, CurrentUser);
    }

    /// <summary>
    /// 获取当前登录用户信息
    /// </summary>
    /// <returns>当前登录用户</returns>
    [HttpGet("userinfo")]
    public async Task<IActionResult> GetUserInfo()
    {
        var user = CurrentUser;
        var detail = await loginService.GetLoginUserInfo(user);
        return Ok(detail);
    }

    [AllowAnonymous]
    [HttpGet("captcha")]
    public async Task<CaptchaResult> CreateCaptcha()
    {
        return await loginService.CreateCaptcha();
    }
}
