using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Security.Claims;
using System.Text.Encodings.Web;

namespace FeiNuo.Core.Security
{
    public class TokenAuthentication : AuthenticationHandler<AuthenticationSchemeOptions>
    {
        public const string AuthenticationScheme = "Bearer";

        private readonly bool isDevelopment;
        private readonly ILogService logService;
        private readonly ITokenService tokenService;
        private readonly SecurityOptions securityOptions = new();

        public TokenAuthentication(ITokenService tokenService, ILogService logService, IConfiguration configuration, IHostEnvironment hostEnvironment, IOptionsMonitor<AuthenticationSchemeOptions> options, ILoggerFactory logger, UrlEncoder encoder) : base(options, logger, encoder)
        {
            this.logService = logService;
            this.tokenService = tokenService;
            isDevelopment = hostEnvironment.IsDevelopment();
            configuration.GetSection(SecurityOptions.ConfigKey).Bind(securityOptions);
        }

        protected override async Task<AuthenticateResult> HandleAuthenticateAsync()
        {
            var token = GetTokenFromRequest();
            if (string.IsNullOrEmpty(token))
            {
                return AuthenticateResult.NoResult();
            }
            try
            {
                var result = isDevelopment && token == AppConstants.SUPER_ADMIN_TOKEN
                    ? new TokenValidationResult(new LoginUser(AppConstants.SUPER_ADMIN, "", [AppConstants.SUPER_ADMIN], []))
                    : await tokenService.ValidateTokenAsync(token);

                if (!result.IsValid)
                {
                    return AuthenticateResult.Fail(result.Message);
                }

                var user = result.LoginUser!;
                Logger.LogDebug("token 认证成功，用户名{Username},token:{token}", user.Username, token);

                var identity = new ClaimsIdentity(user.UserClaims, Scheme.Name, FNClaimTypes.UserName, FNClaimTypes.Role);
                var principal = new ClaimsPrincipal(identity);

                // 保存token
                var properties = new AuthenticationProperties();
                properties.StoreTokens([new AuthenticationToken { Name = "access_token", Value = token }]);

                // 保存用户信息
                var ticket = new AuthenticationTicket(principal, properties, Scheme.Name);

                // 刷新token ，需要前端更新token。
                if (result.RefreshToken)
                {
                    var refreshToken = await tokenService.CreateTokenAsync(user);
                    Response.Headers.Append("fn-refresh-token", refreshToken);
                    //TODO 作废原token? 同时有多个请求时，token作废后，后面的请求会报错
                    //await tokenService.DisableTokenAsync(token);
                }

                return AuthenticateResult.Success(ticket);
            }
            catch (Exception ex)
            {
                Logger.LogError("token 认证失败，token:{token},原因：{Message}", token, ex.Message);
                return AuthenticateResult.Fail(ex.Message);
            }
        }
        protected async override Task HandleForbiddenAsync(AuthenticationProperties properties)
        {
            var msgVo = new MessageResult("没有操作权限", MessageType.Warning)
            {
                Data = $"用户 {Context.User.Identity?.Name} 请求 {Request.Path},没有操作权限"
            };

            // 记录操作日志
            var log = new OperateLog(OperateType.Security, "安全拦截", "没有操作权限") { Success = false };
            log.MergeContextParam(Context);
            await logService.SaveLog(log);

            // 响应403
            Response.StatusCode = StatusCodes.Status403Forbidden;
            await Response.WriteAsJsonAsync(msgVo);
        }

        /// <summary>
        /// 没有登录，或错误的token，返回401
        /// </summary>
        protected override async Task HandleChallengeAsync(AuthenticationProperties properties)
        {
            var msgVo = new MessageResult("请先登录系统", MessageType.Info);
            // 加上错误原因
            var result = await HandleAuthenticateOnceSafeAsync();
            msgVo.Data = result?.Failure?.Message ?? "";

            // 记录操作日志？ 没什么必要。
            //var log = new OperateLog("安全拦截", OperateType.Security, "尚未登录" + msgVo.Detail) { IsSuccess = false };
            //log.MergeContextParam(Context);
            //await logService.SaveLog(log);

            // 响应401
            Response.StatusCode = StatusCodes.Status401Unauthorized;
            await Response.WriteAsJsonAsync(msgVo);
        }

        private string GetTokenFromRequest()
        {
            var token = Request.Headers.Authorization.ToString();
            // 没有认证token
            if (string.IsNullOrWhiteSpace(token))
            {
                return "";
            }
            // 去除前缀
            if (token.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
            {
                token = token["Bearer ".Length..].Trim();
            }
            return token == "Bearer" ? "" : token;
        }
    }
}
