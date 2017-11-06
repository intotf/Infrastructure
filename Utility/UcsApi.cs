using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;

namespace Infrastructure.Utility
{
    /// <summary>
    /// 云之讯客户端
    /// </summary>
    public class UcsUser
    {
        /// <summary>
        /// 账号
        /// </summary>
        public string Account { get; set; }

        /// <summary>
        /// 密码
        /// </summary>
        public string Password { get; set; }
    }



    /// <summary>
    /// 云之讯Api
    /// </summary>
    public static class UcsApi
    {
        private static readonly string Host = "https://api.ucpaas.com";
        private static readonly string SoftVer = "2014-06-30";
        private static readonly string Account = "cb738cfcd6aea9106be9997676e72a02";
        private static readonly string Token = "bc8bbc4ca1f3deb6b3447bf2e711d53f";
        private static readonly string AppId = "d0bf0742fdb34f5187cd9a4dd5c52cac";

        static UcsApi()
        {
            ServicePointManager.ServerCertificateValidationCallback = (a, b, c, d) => true;
        }

        /// <summary>
        /// 申请client帐号
        /// </summary>
        /// <exception cref="clientType">0 开发者计费；1 云平台计费</exception>
        /// <exception cref="charge">充值金额（开发者计费即ClientType为0时，为可选参数）</exception>
        /// <exception cref="mobile">绑定的手机号码。同一个应用内唯一。</exception>
        /// <returns>包体内容</returns>
        public static UcsUser TryCreateClient(int clientType = 0, int charge = 0, string mobile = null)
        {
            try
            {
                return UcsApi.CreateClient(clientType, charge, mobile);
            }
            catch (Exception)
            {
                return default(UcsUser);
            }
        }

        /// <summary>
        /// 申请client帐号
        /// </summary>
        /// <exception cref="clientType">0 开发者计费；1 云平台计费</exception>
        /// <exception cref="charge">充值金额（开发者计费即ClientType为0时，为可选参数）</exception>
        /// <exception cref="mobile">绑定的手机号码。同一个应用内唯一。</exception>
        /// <returns>包体内容</returns>
        public static UcsUser CreateClient(int clientType = 0, int charge = 0, string mobile = null)
        {
            var friendlyName = Guid.NewGuid().ToString().Replace("-", "_");
            var dateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
            var sign = Encryption.GetMD5(Account + Token + dateTime);
            var address = string.Format("{0}/{1}/Accounts/{2}/Clients?sig={3}", Host, SoftVer, Account, sign);
            var auth = Convert.ToBase64String(Encoding.UTF8.GetBytes(Account + ":" + dateTime));

            using (var client = new HttpClient())
            {
                client.Headers.Add("Authorization", auth);
                client.Headers.Add("Content-Type", "application/json");
                client.Headers.Add("Accept", "application/json");
                var body = new { client = new { appId = AppId, friendlyName, clientType, charge, mobile } };
                var bodyJson = JsonSerializer.Serialize(body);

                var json = client.HttpPost(address, bodyJson, Encoding.UTF8);
                var data = JsonSerializer.Deserialize<dynamic>(json);
                if (data.resp.respCode.Value == "000000")
                {
                    var account = (string)data.resp.client.clientNumber.Value;
                    var password = (string)data.resp.client.clientPwd.Value;
                    return new UcsUser { Account = account, Password = password };
                }
                return null;
            }
        }


        /// <summary>
        /// 发送短信
        /// </summary>
        /// <param name="target">短信接收端手机号码</param>
        /// <param name="templateId">短信模板ID</param>
        /// <param name="param">模板内容参数</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="Exception"></exception>
        /// <returns>包体内容</returns>
        public static bool TrySendMsg(string target, string templateId, params string[] param)
        {
            if (target == null)
            {
                throw new ArgumentNullException("target");
            }

            if (templateId == null)
            {
                throw new ArgumentNullException("templateId");
            }

            var dateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
            var sign = Encryption.GetMD5(Account + Token + dateTime);
            var address = string.Format("{0}/{1}/Accounts/{2}/Messages/templateSMS?sig={3}", Host, SoftVer, Account, sign);
            var auth = Convert.ToBase64String(Encoding.UTF8.GetBytes(Account + ":" + dateTime));

            using (var client = new HttpClient())
            {
                client.Headers.Add("Authorization", auth);
                client.Headers.Add("Content-Type", "application/json");
                client.Headers.Add("Accept", "application/json");
                var body = new { templateSMS = new { appId = AppId, templateId = templateId, to = target, param = string.Join(",", param) } };
                var bodyJson = JsonSerializer.Serialize(body);

                try
                {
                    var json = client.HttpPost(address, bodyJson, Encoding.UTF8);
                    var data = JsonSerializer.Deserialize<dynamic>(json);
                    return data.resp.respCode.Value == "000000";
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// 发送手机注册验证码
        /// </summary>
        /// <param name="mobile">手机</param>
        /// <param name="code">验证码</param>
        /// <returns></returns>
        public static bool TrySendRegByMobileMsg(string mobile, string code)
        {
            return UcsApi.TrySendMsg(mobile, "18723", code);
        }
    }
}
