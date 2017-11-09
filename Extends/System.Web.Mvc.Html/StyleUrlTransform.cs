using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Optimization;

namespace System.Web.Mvc
{
    /// <summary>
    /// 样式Url路径绝对化
    /// </summary>
    public class StyleUrlTransform : IItemTransform
    {
        /// <summary>
        /// 虚拟根目录
        /// </summary>
        private string appPath;

        /// <summary>
        /// 样式Url路径绝对化
        /// </summary>
        /// <param name="applicationPath">虚拟根目录</param>
        public StyleUrlTransform(string applicationPath)
        {
            this.appPath = applicationPath;
        }

        /// <summary>
        /// 处理样式
        /// </summary>
        /// <param name="includedVirtualPath">路径路径</param>
        /// <param name="input">内容</param>
        /// <returns></returns>
        public string Process(string includedVirtualPath, string input)
        {
            if (includedVirtualPath == null)
            {
                throw new ArgumentNullException("includedVirtualPath");
            }

            return this.ConvertUrlsToAbsolute(VirtualPathUtility.GetDirectory(includedVirtualPath.Substring(1)), input);
        }

        /// <summary>
        /// 将样式内容包含的url路径转换为绝对路径
        /// </summary>
        /// <param name="baseUrl"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        private string ConvertUrlsToAbsolute(string baseUrl, string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return content;
            }
            Regex regex = new Regex("url\\(['\"]?(?<url>[^)]+?)['\"]?\\)");
            return regex.Replace(content, (MatchEvaluator)(match => ("url(" + this.RebaseUrlToAbsolute(baseUrl, match.Groups["url"].Value) + ")")));
        }

        /// <summary>
        /// 获取路径的绝对路径
        /// </summary>
        /// <param name="baseUrl"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        private string RebaseUrlToAbsolute(string baseUrl, string url)
        {
            if ((string.IsNullOrWhiteSpace(url) || string.IsNullOrWhiteSpace(baseUrl)) || url.StartsWith("/", StringComparison.OrdinalIgnoreCase))
            {
                return url;
            }
            if (!baseUrl.EndsWith("/", StringComparison.OrdinalIgnoreCase))
            {
                baseUrl = baseUrl + "/";
            }

            if (this.appPath != "/")
            {
                baseUrl = this.appPath + baseUrl;
            }
            return VirtualPathUtility.ToAbsolute(baseUrl + url);
        }
    }
}