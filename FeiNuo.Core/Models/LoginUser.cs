﻿using System.ComponentModel;
using System.Security.Claims;
using System.Text.Json.Serialization;

namespace FeiNuo.Core;

/// <summary>
/// 登录用户
/// </summary>
[Serializable]
public class LoginUser
{
    /// <summary>
    /// 用户名
    /// </summary>
    [Description("用户名")]
    public string Username { get; set; } = null!;

    /// <summary>
    /// 姓名/昵称
    /// </summary>
    [Description("姓名/昵称")]
    public string Nickname { get; set; } = string.Empty;

    /// <summary>
    /// 用户密码
    /// </summary>
    [JsonIgnore]
    [Description("用户密码")]
    public string? Password { get; set; }

    /// <summary>
    /// 用户角色
    /// </summary>
    [Description("用户角色")]
    public List<string> Roles { get; set; } = [];

    /// <summary>
    /// 权限字符串
    /// </summary>
    [Description("权限字符串")]
    public List<string> Permissions { get; set; } = [];

    /// <summary>
    /// 其他信息
    /// </summary>
    [Description("其他信息")]
    public string UserData { get; set; } = string.Empty;

    /// <summary>
    /// 请求的客户端信息，方便传参用，有需要时需手动在Controller中赋值 
    /// </summary>
    [Description("请求的客户端信息，方便传参用，有需要时需手动在Controller中赋值 ")]
    public RequestClient? RequestClient { get; set; }

    /// <summary>
    /// 默认构造函数
    /// </summary>
    public LoginUser() { }

    /// <summary>
    /// 根据User.Claims构造用户对象
    /// </summary>
    /// <param name="claims"></param>
    public LoginUser(IEnumerable<Claim> claims) { UserClaims = claims; }

    /// <summary>
    /// 构造函数
    /// roles和permissions字符串中不能有英文逗号
    /// </summary>
    public LoginUser(string username, string nickname, string password, IEnumerable<string> roles, IEnumerable<string>? permissions = null, string? userData = null)
    {
        Username = username;
        Nickname = nickname;
        Password = password;
        Roles = [.. roles];
        if (permissions != null) Permissions = [.. permissions];
        if (!string.IsNullOrWhiteSpace(userData)) UserData = userData;
    }

    /// <summary>
    /// 是否是超级管理员
    /// </summary>
    public bool IsSuperAdmin
    {
        get
        {
            return Username == AppConstants.SUPER_ADMIN || Roles.Any(t => t == AppConstants.SUPER_ADMIN);
        }
    }

    /// <summary>
    /// 是否有角色
    /// </summary>
    public bool HasRole(string role)
    {
        return Roles.Any(a => a == role);
    }

    /// <summary>
    /// 用户信息和User.Identify.Claims的转换
    /// </summary>
    [JsonIgnore]
    public IEnumerable<Claim> UserClaims
    {
        get
        {
            var claims = new List<Claim>()
            {
                new (FNClaimTypes.UserName, Username),
            };
            if (!string.IsNullOrWhiteSpace(Nickname))
            {
                claims.Add(new(FNClaimTypes.NickName, Nickname));
            }
            foreach (var role in Roles)
            {
                claims.Add(new(FNClaimTypes.Role, role));
            }
            foreach (var perm in Permissions)
            {
                claims.Add(new(FNClaimTypes.Permission, perm));
            }
            if (!string.IsNullOrEmpty(UserData))
            {
                claims.Add(new(FNClaimTypes.Data, UserData));
            }
            return claims;
        }
        private set
        {
            if (!value.Any()) return;
            Username = value.SingleOrDefault(a => a.Type == FNClaimTypes.UserName)!.Value;
            Nickname = value.SingleOrDefault(a => a.Type == FNClaimTypes.NickName)?.Value ?? string.Empty;

            Roles = [.. value.Where(a => a.Type == FNClaimTypes.Role || a.Type == ClaimTypes.Role).Select(a => a.Value).Distinct()];
            Permissions = [.. value.Where(a => a.Type == FNClaimTypes.Permission).Select(a => a.Value).Distinct()];

            UserData = value.SingleOrDefault(a => a.Type == FNClaimTypes.Data)?.Value ?? string.Empty;
        }
    }
}

public class FNClaimTypes
{
    public const string UserId = "uid";
    public const string UserName = "user";
    public const string NickName = "name";
    public const string Role = "role";
    public const string Permission = "perm";
    public const string Data = "data";
    public const string NotBefore = "nbf";
    public const string IssuedAt = "iat";
    public const string Expire = "exp";
}
