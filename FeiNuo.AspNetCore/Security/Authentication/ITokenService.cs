namespace FeiNuo.AspNetCore.Security.Authentication;

/// <summary>
/// Token服务类
/// </summary>
public interface ITokenService
{
    /// <summary>
    /// 根据用户信息创建Token
    /// </summary>
    Task<string> CreateTokenAsync(LoginUser user);

    /// <summary>
    /// 验证Token合法性，通过后根据token获取用户信息,不通过写入原因
    /// </summary>
    Task<TokenValidationResult> ValidateTokenAsync(string token);

    /// <summary>
    /// 作废Token
    /// </summary>
    Task DisableTokenAsync(string token);
}