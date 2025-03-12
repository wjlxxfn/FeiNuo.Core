using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FeiNuo.AspNetCore.Security
{
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
}
