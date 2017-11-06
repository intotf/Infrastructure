using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Routing;

namespace System.Web.Mvc
{
    /// <summary>
    /// 版本路径映射
    /// </summary>
    public static class VersionRouting
    {
        /// <summary>
        /// Api版本匹配路由
        /// 将Url的{version}标记追加到Action名之后，两者中间下划线分开
        /// 此路由规则需要优先默认路由规则添加到路由表中
        /// </summary>
        /// <param name="routes">路由列表</param>
        /// <param name="handler">处理者</param>        
        /// <param name="name">名称</param>
        /// <param name="url">URL，要包含{version}标志</param>
        /// <param name="defaults">初始值</param>
        /// <param name="namespaces">空间</param>
        public static void MapRoute(this RouteCollection routes, IRouteHandler handler, string name, string url, object defaults, string[] namespaces)
        {
            var route = new VersionRoute(url, handler)
            {
                Defaults = CreateRouteValueDictionary(defaults),
                Constraints = new RouteValueDictionary(),
                DataTokens = new RouteValueDictionary()
            };

            if ((namespaces != null) && (namespaces.Length > 0))
            {
                route.DataTokens["Namespaces"] = namespaces;
            }
            routes.Add(name, route);
        }

        /// <summary>
        /// 创建路由数据字典
        /// </summary>
        /// <param name="values">数据</param>
        /// <returns></returns>
        private static RouteValueDictionary CreateRouteValueDictionary(object values)
        {
            var dictionary = values as IDictionary<string, object>;
            if (dictionary != null)
            {
                return new RouteValueDictionary(dictionary);
            }
            return new RouteValueDictionary(values);
        }

        /// <summary>
        /// API版本路由规则
        /// </summary>
        public class VersionRoute : Route
        {
            public VersionRoute(string url, IRouteHandler handler)
                : base(url, handler)
            {
            }

            /// <summary>
            /// 获取路由数据
            /// </summary>
            /// <param name="httpContext"></param>
            /// <returns></returns>
            public override RouteData GetRouteData(HttpContextBase httpContext)
            {
                var routeData = base.GetRouteData(httpContext);
                if (routeData == null)
                {
                    return null;
                }

                var version = routeData.Values["version"].ToString();
                if (version.ToLower() == "v1") //V1略过
                {
                    return routeData;
                }
                else if (Regex.IsMatch(version, @"^v\d*$", RegexOptions.IgnoreCase) == false)
                {
                    return null;
                }

                var routeValues = routeData.Values.ToArray();
                routeData.Values.Clear();
                foreach (var kv in routeValues)
                {
                    if (kv.Key == "action")
                    {
                        routeData.Values.Add(kv.Key, string.Format("{0}_{1}", kv.Value, version));
                    }
                    else
                    {
                        routeData.Values.Add(kv.Key, kv.Value);
                    }
                }
                return routeData;
            }
        }
    }
}
