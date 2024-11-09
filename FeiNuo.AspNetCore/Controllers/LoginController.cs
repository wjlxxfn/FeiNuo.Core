using FeiNuo.AspNetCore.Security.FormLogin;
using FeiNuo.Core.Utilities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

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
    [HttpPost("login")]
    public async Task<string> HandleLogin([FromBody] LoginForm loginForm)
    {
        var log = new OperateLog(OperateType.Login, "用户登录", "");
        var stopwatch = Stopwatch.StartNew();
        try
        {
            var token = await loginService.HandleLogin(loginForm);
            log.LogContent = token;
            return token;
        }
        catch (Exception ex)
        {
            log.Success = false;
            log.LogContent = ex.Message;
            throw;
        }
        finally
        {
            stopwatch.Stop();
            log.ExecuteTime = stopwatch.ElapsedMilliseconds;
            log.MergeContextParam(HttpContext);
            log.RequestParam = JsonUtils.Serialize(loginForm);
            log.OperateBy = loginForm.Username;
            await logService.SaveLog(log);
        }
    }

    /// <summary>
    /// 退出登录
    /// </summary>
    [AllowAnonymous]
    [HttpPost("logout")]
    public async void HandleLogout()
    {
        var log = new OperateLog(OperateType.Logout, "退出登录", "");
        var stopwatch = Stopwatch.StartNew();
        try
        {
            var token = await GetAccessToken();
            if (string.IsNullOrWhiteSpace(token) || User == null)
            {
                log.LogContent = "没有登录token";
                log.Success = false;
                return;
            }
            log.OperateBy = CurrentUser.Username;
            log.RequestParam = token;
            await loginService.HandleLogout(token, CurrentUser);

        }
        catch (Exception ex)
        {
            log.Success = false;
            log.LogContent = ex.Message;
            throw;
        }
        finally
        {
            stopwatch.Stop();
            log.ExecuteTime = stopwatch.ElapsedMilliseconds;
            log.MergeContextParam(HttpContext);
            await logService.SaveLog(log);
        }
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
