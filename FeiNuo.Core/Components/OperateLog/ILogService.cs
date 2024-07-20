﻿using FeiNuo.Core.Utilities;

namespace FeiNuo.Core
{
    /// <summary>
    /// 日志保存服务，需调用方自己实现
    /// </summary>
    public interface ILogService
    {
        /// <summary>
        /// 保存日志
        /// </summary>
        public Task SaveLog(OperateLog log);

        /// <summary>
        /// 保存日志
        /// </summary>
        public Task SaveLog(OperateType operateType, string logTitle, string logDetail);

        /// <summary>
        /// 保存日志
        /// </summary>
        public async Task SaveLog(OperateType operateType, string logTitle, dynamic logDetail)
        {
            var detail = JsonUtils.Serialize(logDetail);
            await SaveLog(operateType, logTitle, detail);
        }
    }
}