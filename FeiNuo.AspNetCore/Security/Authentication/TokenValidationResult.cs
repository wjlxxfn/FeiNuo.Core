namespace FeiNuo.AspNetCore.Security.Authentication;

/// <summary>
/// token校验结果
/// </summary>
public class TokenValidationResult
{
    /// <summary>
    /// 是否验证通过
    /// </summary>
    public bool IsValid { get; private set; }

    /// <summary>
    /// 验证失败的原因
    /// </summary>
    public string Message { get; private set; } = string.Empty;

    /// <summary>
    /// 验证成功后转成LoginUser对象
    /// </summary>
    public LoginUser? LoginUser { get; private set; }

    /// <summary>
    /// 是否需要刷新token
    /// </summary>
    public bool RefreshToken { get; set; } = false;

    /// <summary>
    /// 验证失败的构造函数，传入失败原因
    /// </summary>
    public TokenValidationResult(string message)
    {
        IsValid = false;
        Message = message;
    }

    /// <summary>
    /// 验证成功的构造函数，传入转换后的用户对象
    /// </summary>
    public TokenValidationResult(LoginUser user)
    {
        IsValid = true;
        LoginUser = user;
    }
}
