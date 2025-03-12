# FeiNuo.Core
## 功能介绍
    该类库基于Net8开发，提供NetCore开发常用功能的封装。
1. 提供常用数据模型；
2. 提供常用工具类和扩展功能：如JsonUtil,StringExtensions等;
3. 提供常用组件的封装，如：Excel操作，验证码生成等；

## 一、常用数据模型  
1. 实体类基类，提供操作人相关字段
```
   1、BaseDto: CreateBy,CreateTime
   2、BaseEntity:  CreateBy,CreateTime,UpdateBy,UpdateTime
```
2. 分页模型
```
   1、PageResult:  TotalCount(int), DataList(IEnumerable<TEntity>)
   2、Pager: PageNo(int), PageSize(int), SortBy(List<SortItem>)
   3、SortItem: SortField(string), SortType(SortTypeEnum)
   4、SortTypeEnum: ASC,DESC
```   
3. 查询模型
```
   BaseQuery: 查询基类，提供一个Search模糊搜索条件和时间范围字段StartDate/EndDate
   AbstractQuery: 查询条件抽象类，提借查询条件的封装拼成Linq表达式，方便EF直接调用
```   
4. 登录用户LoginUser
```
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

                Roles = value.Where(a => a.Type == FNClaimTypes.Role || a.Type == ClaimTypes.Role).Select(a => a.Value).Distinct().ToList();
                Permissions = value.Where(a => a.Type == FNClaimTypes.Permission).Select(a => a.Value).Distinct().ToList();

                UserData = value.SingleOrDefault(a => a.Type == FNClaimTypes.Data)?.Value ?? string.Empty;
            }
        }
     }
```

5. 其他模型
```
    1、UpdateDto: 通用更新模型,适用于更新少量常见字段，方便传参用
    2、SelectOption: 前端常用的下拉列表选项，提供Lable,Value,Disabled,Color,Data常用字段
    3、TreeOption: 继承自SelectOption，添加Children字段，形成树形结构
    4、RouteVO，MetaVO,MenuVO: Vue路由或菜单对应的实体对象
    5、RequestClient: 封装请求客户端的信息：Ip,Os,Browser,Device
    6、MessageResult: 用于前后端传消息用：Type,Message,Data,Timestamp

```

## 二、提供常用工具类和扩展功能
1. JsonUtils: 使用System.Text.Json处理序列化，默认添加以下序列化配置
```
    public static void MergeSerializerOptions(JsonSerializerOptions options)
    {
        // 空值不输出
        options.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull;
        // 阻断循环引用
        options.ReferenceHandler = ReferenceHandler.IgnoreCycles;
        // 设置支持的编码,防止转义
        options.Encoder = JavaScriptEncoder.Create(UnicodeRanges.All);
        // 小驼峰转换
        options.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
        // 支持从字符转成数字
        options.NumberHandling = JsonNumberHandling.AllowReadingFromString;
        // JSON格式化配置，日期格式 yyyy-MM-dd HH:mm:ss
        options.Converters.Add(new DateTimeConverter());
    }
```
2. DateTimeUtils: 日期相关，目前只定义了两个常用日期常量，其他待添加
  ```
    /// <summary>
    /// SQL Server中零日期值，主要用于日期字段的默认值。
    /// </summary>
    public static readonly DateTime ZeroSqlDateTime = new(1900, 1, 1);

    /// <summary>
    /// SQL Server中最大日期值，主要用于查询条件：结束日期
    /// </summary>
    public static readonly DateTime MaxSqlDateTime = new(9999, 1, 1);
  ```
3. 常用类的扩展方法
```
    StringExtensions:
    EnumExtensions:
    DateTimeExtensions:
    EnumExtensions:
    ExpressionExtensions:
    QueryableExtensions:
```

## 三、提供常用组件功能
   _部分功能需引入相应的依赖库_
#### 1、对NPOI Excel的功能封装
  > 该功能需要此入 NPOI (>=2.7.2)

#### 2、验证码功能封装
  > 该功能需要此入 SixLabors.ImageSharp.Drawing (>=2.1.5) 

   系统提供CaptchaUtils工具类直接生成验证码；另外提供配置选项类CaptchaOptions。

  *生成的验证码包括文本和图片的base64格式，后续使用需自己实现*
  
#### 2、操作日志定义
  1、OperateLog  定义操作日志
  2、OperateType 定义操作类型
  3、ILogService 定义操作接口：保存日志用
