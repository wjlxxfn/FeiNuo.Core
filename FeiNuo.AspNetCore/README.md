# FeiNuo.AspNetCore
## ���ܽ���
    ����Ŀ����Net8���������ڸ�������WebApi��Ŀ��    
1. ��װ�Զ�ע�룬�쳣������¼��֤��ϵͳ��Ȩ��ͨ�ô��룻
2. �ṩ���ù��������չ���ܣ���JsonUtil,StringExtensions��;
3. �ṩ��������ķ�װ���磺Excel��������־������,��֤�����ɵȣ�

## һ���Զ�ע��
    builder.Services.AddAppServices();
1. ϵͳ�Զ�ע��ʵ��IService�ӿڻ�̳�BaseService���������
    1. ���������ʵ�ֵĽӿ�(��IService)����ݽӿ���ע��
    2. ���û��ʵ�ֽӿ�(��IService)������ݵ�ǰ����ע��
    3. ע�����������ȫ��Ϊ `ServiceLifetime.Scoped`
    4. ���Ĭ��ע��Ĳ����ʣ������[Service]�����Զ���
    
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
| /login    | POST  | LoginForm  | Token����Ϊ��
| /logout   | POST  | ��         | ��
| /userinfo | GET   | ��         | LoginUser
| /captcha  | GET   | ��         | CaptchaResult  

#### ����ʵ�� ILoginService�����ṩ��¼������û���Ϣ
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
#### ��ʵ��ILoginService��̳�LoginService������ط������޸�Ĭ�ϵĵ�¼�ӿ�ʵ���߼�
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
/// <summary>
    /// Token������
    /// </summary>
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

## ���ҳ���

## PoiExcel��װ

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
