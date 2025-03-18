using Microsoft.Extensions.Logging;

namespace FeiNuo.Core.Login;

public interface ILoginUserService
{
    /// <summary>
    /// 通过用户名查询用户信息：包括，用户名，密码，角色，权限
    /// </summary>
    Task<LoginUser?> LoadUserByUsername(string username);

    /// <summary>
    /// 获取用户信息，根据前端需要返回
    /// </summary>
    virtual Task<Dictionary<string, object>> GetLoginUserInfo(LoginUser user)
    {
        var map = new Dictionary<string, object>
        {
            { "username", user.Username },
            { "roles", user.Roles },
            { "permissions", user.Permissions },
            { "data", user.UserData??"" },
        };
        if (!string.IsNullOrWhiteSpace(user.Nickname))
        {
            map.Add("nickname", user.Nickname);
        }
        return Task.FromResult(map);
    }

    /// <summary>
    /// 验证登录密码，默认直接通过密码明文对比,有加密的可重写该方法
    /// </summary>
    /// <param name="form">用户填写的登录表单数据</param>
    /// <param name="user">数据库中的用户信息</param>
    /// <returns>验证结果</returns>
    virtual bool ValidatePassword(LoginForm form, LoginUser user)
    {
        return form.Password == user.Password;
    }
}



public class SimpleLoginUserService : ILoginUserService
{
    public SimpleLoginUserService(ILogger<SimpleLoginUserService> logger)
    {
        logger.LogError("未注入查询用户的服务类。如需使用登录接口,需实现接口ILoginUserService并注入容器。");
    }

    public Task<LoginUser?> LoadUserByUsername(string username)
    {
        throw new NotImplementedException();
    }
}
