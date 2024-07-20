using System.Security.Claims;
using System.Text.Json.Serialization;

namespace FeiNuo.Core
{
    /// <summary>
    /// 登录用户
    /// </summary>
    [Serializable]
    public class LoginUser
    {
        /// <summary>
        /// 用户名
        /// </summary>
        public string Username { get; set; } = null!;

        /// <summary>
        /// 用户密码
        /// </summary>
        [JsonIgnore]
        public string? Password { get; set; }

        /// <summary>
        /// 用户角色
        /// </summary>
        public List<string> Roles { get; set; } = [];

        /// <summary>
        /// 权限字符串
        /// </summary>
        public List<string> Permissions { get; set; } = [];

        /// <summary>
        /// 其他信息
        /// </summary>
        public string UserData { get; set; } = string.Empty;

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
        public LoginUser(string username, string password, IEnumerable<string> roles, IEnumerable<string>? permissions = null, string? userData = null)
        {
            Username = username;
            Password = password;
            Roles = roles.ToList();
            if (permissions != null) Permissions = permissions.ToList();
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
                    new (FNClaimTypes.Name, Username),
                    // new (FNClaimTypes.NotBefore, LoginTime.ToUnixTimeSeconds().ToString(),ClaimValueTypes.Double),
                };
                if (Roles.Count > 0)
                {
                    claims.Add(new(FNClaimTypes.Role, string.Join(",", Roles)));
                }

                if (Permissions.Count > 0)
                {
                    claims.Add(new(FNClaimTypes.Permission, string.Join(",", Permissions)));
                }

                if (!string.IsNullOrEmpty(UserData))
                {
                    claims.Add(new(FNClaimTypes.Data, UserData));
                }

                return claims;
            }
            private set
            {
                Username = value.SingleOrDefault(a => a.Type == FNClaimTypes.Name)!.Value;

                var roles = value.Where(a => a.Type == FNClaimTypes.Role).SingleOrDefault();
                Roles = null == roles ? [] : [.. roles.Value.Split(',')];

                var perms = value.Where(a => a.Type == FNClaimTypes.Permission).SingleOrDefault();
                Permissions = null == perms ? [] : [.. perms.Value.Split(',')];

                UserData = value.SingleOrDefault(a => a.Type == FNClaimTypes.Data)?.Value ?? "";

                //var nbf = long.Parse(value.SingleOrDefault(a => a.Type == FNClaimTypes.NotBefore)?.Value ?? "0");
                //LoginTime = DateTimeOffset.FromUnixTimeSeconds(nbf).UtcDateTime.ToLocalTime();
            }
        }
    }

    internal class FNClaimTypes
    {
        public const string UserId = "uid";
        public const string Name = "name";
        public const string Role = "role";
        public const string Permission = "perm";
        public const string Data = "data";
        public const string NotBefore = "nbf";
        public const string IssuedAt = "iat";
        public const string Expire = "exp";
    }
}
