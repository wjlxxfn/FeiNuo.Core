namespace FeiNuo.Core;

/// <summary>
/// 实体 基类
/// </summary>
public class BaseEntity : ICloneable
{
    /// <summary>
    /// 创建人
    /// </summary>
    public string? CreateBy { get; set; }

    /// <summary>
    /// 创建时间
    /// </summary>
    public DateTime? CreateTime { get; set; }

    /// <summary>
    /// 修改人
    /// </summary>
    public string? UpdateBy { get; set; }

    /// <summary>
    /// 修改时间
    /// </summary>
    public DateTime? UpdateTime { get; set; }


    /// <summary>
    /// 添加操作人信息
    /// </summary>
    public void AddOperator(string user, bool isNew)
    {
        if (isNew) InitCreate(user);
        else InitUpdate(user);
    }

    /// <summary>
    /// 初始化创建人和创建时间
    /// </summary>
    public void InitCreate(string user)
    {
        CreateBy = user;
        CreateTime = DateTime.Now;
        InitUpdate(user);
    }

    /// <summary>
    /// 初始化修改人和修改时间
    /// </summary>
    public void InitUpdate(string user)
    {
        UpdateBy = user;
        UpdateTime = DateTime.Now;
    }

    /// <summary>
    /// 浅拷贝，调用Object.MemberwiseClone实现，如需深拷贝需另外实现
    /// </summary>
    public object Clone()
    {
        return MemberwiseClone();
    }
}
