namespace FeiNuo.Core;

/// <summary>
/// 登录接口
/// </summary>
public interface ILoginService
{
    /// <summary>
    /// 登录系统
    /// </summary>
    Task<string> HandleLogin(LoginForm form);

    /// <summary>
    /// 退出登录
    /// </summary>
    /// <param name="token">要退出的token</param>
    /// <param name="user">当前操作用户</param>
    Task HandleLogout(string token, LoginUser user);

    /// <summary>
    /// 获取登录用户的详细信息，前端用啥，就返回啥，默认返回LoginUser里的信息
    /// </summary>
    Task<Dictionary<string, object>> GetLoginUserInfo(LoginUser user)
    {
        var map = new Dictionary<string, object>
        {
            { "username", user.Username },
            { "roles", user.Roles },
            { "permissions", user.Permissions },
        };
        return Task.FromResult(map);
    }

    /// <summary>
    /// 生成验证码
    /// </summary>
    /// <returns></returns>
    Task<CaptchaResult> CreateCaptcha()
    {
        return Task.FromResult(new CaptchaResult());
    }
}
