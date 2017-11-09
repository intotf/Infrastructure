using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace Infrastructure.Utility
{
    public static class HttpHelper
    {
        /// <summary>
        /// 查询客户端ip
        /// </summary>
        /// <returns></returns>
        public static string GetClientIp()
        {
            HttpRequest objRequest = HttpContext.Current.Request;
            return (objRequest.ServerVariables["HTTP_VIA"] != null) ? objRequest.ServerVariables["HTTP_X_FORWARDED_FOR"] : objRequest.ServerVariables["REMOTE_ADDR"];
        }

        /// <summary>
        /// 获取相对路径的绝对路径
        /// </summary>
        /// <param name="url">相对路径</param>
        /// <returns>绝对路径</returns>
        public static string GetAbsolutePath(string url)
        {
            Uri uri;
            if (Uri.TryCreate(HttpContext.Current.Request.Url, url, out uri))
            {
                return uri.ToString();
            }
            return url;
        }
    }
}
