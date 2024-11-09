# FeiNuo.AspNetCore
## ���ܽ���
    ����Ŀ����Net8���������ڸ�������WebApi��Ŀ��    
1. ��װ�Զ�ע�룬�쳣������¼��֤��ϵͳ��Ȩ��ͨ�ù��ܣ�
2. �ṩ���ù��������չ���ܣ���JsonUtil,StringExtensions��;
3. �ṩ��������ķ�װ���磺Excel��������־������, ��֤�����ɵȣ�

## һ���Զ�ע��
    builder.Services.AddAppServices();
1. ϵͳ�Զ�ע��ʵ��IService�ӿڻ�̳�BaseService���������
    1. ���������ʵ�ֵĽӿ�(��IService)����ݽӿ���ע��
    2. ���û��ʵ�ֽӿ�(��IService)������ݵ�ǰ����ע��
    3. ע�����������ȫ��Ϊ `ServiceLifetime.Scoped`
    4. ���Ĭ��ע��Ĳ����ʣ������[Service]�����Զ���
    5. �����[Service]���ԣ����������[Service]���ԵĹ���ע��
    
2. ϵͳ�Զ�ע���ע��[Service]���Ե�������
    1. �����������ע�����ͣ�����������ָ��������Ϊ׼
    2. ���������û��ע�����ͣ�����ݵ�ǰ��ʵ�ֵĽӿ�ע��
    3. �����ǰ��Ҳû��ʵ�ֽӿڣ�����ݵ�ǰ����ע��
    4. ��������Ĭ��Ϊ `ServiceLifetime.Scoped`���������������޸�
    
3.  Ĭ��ע���ڴ滺����ڴ�ֲ�ʽ����

## ����API��Ϊ����
    builder.Services.AddAppControllers();
1. ϵͳԼ��ͳһ����HTTP״̬�뷵������
2. ֧��DateOnly���ͽ��մ�ʱ�������
3. ������Ӧ������״̬��2XX,ֱ�ӷ���ϵͳ����
4. �쳣��Ӧ��
    1. û�е�¼��Ϣ������ 401,UnauthorizedResult
    2. û��Ȩ�ޣ����� 403,ForbidResult
    3. ģ��������֤��ͨ�������� 400,BadRequestResult
    4. ϵͳ�Զ����쳣��Ϣ������ 422,UnprocessableEntityResult
    5. ϵͳ�쳣������ 500
1. ͳһ���л�����
    1. nullֵĬ�ϲ����
    1. ���ѭ������
    1. ����ת�����ԣ�С�շ�
    1. ֧�ִ��ַ���ת������
    1. ��������Ĭ�������ʽ yyyy-MM-dd HH:mm:ss
    
## ������֤��Ȩ
    builder.Services.AddAppSecurity(builder.Configuration);

### 1����¼�ӿ�
    �ṩ��¼���Rest�ӿ�  
    ��Ҫ��Ҫע���ѯ��¼�û���ʵ���ࣺILoginUserService  

|    �ӿ�   | ����  |     ����   |   ����  |
|    ----   | :---: |    ------  |   ----  |
| /login    | POST  | LoginForm  | Token�ַ���
| /logout   | POST  | ��         | ��
| /userinfo | GET   | ��         | LoginUser
| /captcha  | GET   | ��         | CaptchaResult  

#### ����ʵ�� ILoginUserService�����ṩ��¼������û���Ϣ
```
public interface ILoginUserService
{
    /// <summary>
    /// ͨ���û�����ѯ�û���Ϣ���������û��������룬��ɫ��Ȩ��
    /// </summary>
    Task<LoginUser?> LoadUserByUsername(string username);

    /// <summary>
    /// ��ȡ�û���Ϣ������ǰ����Ҫ����
    /// </summary>
    virtual Task<Dictionary<string, object>> GetLoginUserInfo(LoginUser user)
    {
        var map = new Dictionary<string, object>
        {
            { "username", user.Username },
            { "roles", user.Roles },
            { "permissions", user.Permissions },
            { "data", user.UserData??"" },
        };
        if (!string.IsNullOrWhiteSpace(user.Nickname))
        {
            map.Add("nickname", user.Nickname);
        }
        return Task.FromResult(map);
    }

    /// <summary>
    /// ��֤��¼���룬Ĭ��ֱ��ͨ���������ĶԱ�,�м��ܵĿ���д�÷���
    /// </summary>
    /// <param name="form">�û���д�ĵ�¼������</param>
    /// <param name="user">���ݿ��е��û���Ϣ</param>
    /// <returns>��֤���</returns>
    virtual bool ValidatePassword(LoginForm form, LoginUser user)
    {
        return form.Password == user.Password;
    }
}
```
#### ��ʵ��ILoginService��̳�LoginService��д��ط������޸�Ĭ�ϵĵ�¼�ӿ�ʵ���߼�
```
public interface ILoginService
{
    /// <summary>
    /// ��¼ϵͳ
    /// </summary>
    Task<string> HandleLogin(LoginForm form);

    /// <summary>
    /// �˳���¼
    /// </summary>
    /// <param name="token">Ҫ�˳���token</param>
    /// <param name="user">��ǰ�����û�</param>
    Task HandleLogout(string token, LoginUser user);

    /// <summary>
    /// ��ȡ��¼�û�����ϸ��Ϣ��ǰ����ɶ���ͷ���ɶ��Ĭ�Ϸ���LoginUser�����Ϣ
    /// </summary>
    Task<Dictionary<string, object>> GetLoginUserInfo(LoginUser user)
    {
        var map = new Dictionary<string, object>
        {
            { "username", user.Username },
            { "roles", user.Roles },
            { "permissions", user.Permissions },
        };
        return Task.FromResult(map);
    }

    /// <summary>
    /// ������֤��
    /// </summary>
    /// <returns></returns>
    Task<CaptchaResult> CreateCaptcha()
    {
        return Task.FromResult(new CaptchaResult());
    }
}
```
### 2����֤ģ��
#### ����token��֤ģ�飬ϵͳĬ��ʵ��Jwt�ͻ���Cache��token
> Ҳ��ʵ��ITokenService�ӿڣ��Զ���token��ʵ�ַ�ʽ��Ȼ��ʵ����ע��ϵͳ����
```
public interface ITokenService
{
    /// <summary>
    /// �����û���Ϣ����Token
    /// </summary>
    Task<string> CreateTokenAsync(LoginUser user);

    /// <summary>
    /// ��֤Token�Ϸ��ԣ�ͨ�������token��ȡ�û���Ϣ,��ͨ��д��ԭ��
    /// </summary>
    Task<TokenValidationResult> ValidateTokenAsync(string token);

    /// <summary>
    /// ����Token
    /// </summary>
    Task DisableTokenAsync(string token);
}
```

### 3����Ȩģ��
 1. ��������Ա[LoginUser.IsSuperAdmin]�û�ӵ������Ȩ��
 2. �ṩ����[Permission]���ɸ����û�Ȩ��[LoginUser.Permissions]��֤

## �ġ���־ģ��

## �塢��ҳ���

## ����PoiExcel��װ
    ���������Է�����װ��ExcelConfig������У�Ȼ��ʹ��PoiHelper��װPOIʵ��
    �ṩExcelExporter��ExcelImporter����Excel�ĵ��뵼��
1. POIUtils�ṩ���õ�Excel����
1. ExcelConfig��ExcelSheet,ExcelColumn,ExcelStyle ������Excel���԰������,�������壬��POIû��ϵ 
1. POIHelper��ʹ��POI�ͽ�ExcelConfig������Excel����
1. ExcelExporter��ʵ�ֿ��ٵ������ݣ��ڲ�Ҳ�Ƿ�װExcelConfigʵ�֡�ʹ�ü�����
    1. �����к����ݶ�Ӧ��ϵ
    1. ֧�������п���ʽ
    1. ֧�ֶ��б��⣬������ʽ���õ�
```
[HttpGet("export")]
public IActionResult Export()
{
    var lstData = new List<UserEntity>();
    var excel = new ExcelExporter("�û����ݵ���.xlsx")     // ���嵼���ļ���
        .AddDataSheet(lstData, [                            // ���Sheet,�ɶ����ӣ�ͬʱ�������Sheet
            new("����", s => s.Username, 20, "@"),         // ��ʽ���ã��п�20���ı���ʽ
            new("�ձ�", s => s.UserData?.Gender, 15),      // ���ȡ��������������ʾ�ļ���
            new("����#ʡ",s => s.Addr.Province,15),        // ���б��⣺��#�ָ����Զ��ϲ�
            new("����#��",s => s.Addr.City),
            new("ѧ��",s => s.Education)
         ], s =>                                           // ���������ø�������
        {
            s.Description = "˵������";
            s.DescriptionStyle.FontBold = true;
            s.MainTitle = "������";
            s.MainTitleColSpan = 15;
        });
    return File(excel.GetBytes(), excel.ContentType, excel.FileName);
}
```
## �ߡ��������ɣ�ʹ��EFCore Tools�����򹤳�    
  ��������˵����efpt.renaming.json
  1. ȥ����ͷ����Ӻ�����������ʽ����
```
[
  {
    "SchemaName": "dbo",
    "UseSchemaName": false,
    "TableRegexPattern": "(^md|sys)_(?<table>.+$)",
    "TablePatternReplaceWith": "${table}_entity",
    "Tables": []
  }
]
```
  2. ��Ҫע��sqlserver����ҪSchemaName,mysql������SchemaName��һ��
  
  3. ʹ��Ǩ��˵��Migration
        
    ��Ҫ��WebApi��Ŀ��� Microsoft.EntityFrameworkCore.Tools
    �ڳ��������������̨ʹ������   Add-Migration���������ݿ�ʹ�� Update-Database
    �磺  Add-Migration InitialCreate
    ��������Ŀʱ����Ҫ���� Assembly
    �磺opt.UseSqlServer(conn, b => b.MigrationsAssembly("WebApi"));


## ������չ�͹�����

## ϵͳ����������
��������ֵΪϵͳĬ��ֵ��ֻ������Ҫ�����Ĳ��ּ��ɡ�
 ```
{
  "FeiNuo": {
    "AppCode": "FeiNuo",        // ϵͳ����
    "AppName": "��ŵ",          // ϵͳ����
    "AppVersion": "1.0.0",      // �汾��

    // ��¼��֤�������
    "Security": {
      "TokenType": "Jwt",      // Token���ͣ�Jwt/Cache
      "TokenExpiration": 7200, // Token��ʱʱ�䣬��λ�룬0����ʱ

      // Jwt����
      "Jwt": {
        "SigningKey": "abcdefg1234", // ����ǩ��Key
        "ClockSkew": 1800,       // JWT�Ļ���ʱ�䣬��λ�룬ʵ�ʵĹ���ʱ�� = exp+clockskew
        "Issuer": "",            // ֤��䷢�ߣ�������Ĭ�ϲ���֤
        "Audience": ""           // ���ڷ���������Ĭ�ϲ���֤
      },

      // ��֤������
      "Captcha": {
        "Enabled": true,                // �Ƿ�����
        "Timeout": 300,                 // ��ʱʱ��,��λ��
        "Length": 4,                    // ��֤���ַ�����
        "Width": 100,                   // ͼƬ��ȣ���λpx
        "Height": 40,                   // ͼƬ�߶ȣ���λpx
        "LineCount": 4,                 // ����������
        "ChaosDensity": 40,             // ����ܶȣ���������ػ�/�ܶȼ������������0��ʾ������
        "FontSize": 30,                 // �����С��λ��
        "BackGroundColor": "#ffffff",   // ͼƬ����ɫ
      }
    }
  },
}
 ```
