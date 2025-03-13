namespace FeiNuo.Core;

/// <summary>
/// 服务类接口
/// </summary>
public interface BaseService
{

}

/// <summary>
/// 服务类基类：指定数据库上下文和实体类型，提供基础操作
/// </summary>
public abstract class BaseService<T> : BaseService where T : class, new()
{

}