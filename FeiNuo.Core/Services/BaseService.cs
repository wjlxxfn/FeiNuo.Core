namespace FeiNuo.Core;

/// <summary>
/// 服务类接口
/// </summary>
#pragma warning disable IDE1006 // 命名样式
public interface BaseService
#pragma warning restore IDE1006 // 命名样式
{

}

/// <summary>
/// 服务类基类：指定数据库上下文和实体类型，提供基础操作
/// </summary>
public abstract class BaseService<T> : BaseService where T : class, new()
{

}