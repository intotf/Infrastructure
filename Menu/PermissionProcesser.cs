using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using System.Web.Routing;

namespace Infrastructure.Menu
{
    /// <summary>
    /// 权限处理类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public static class PermissionProcesser
    {
        /// <summary>
        /// 处理权限
        /// 返回当前节点
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="routeData">路由数据</param>
        /// <param name="menuNode">菜单节点</param>
        /// <param name="failureAction">权限验证不通过</param>
        public static MenuNode<T> Process<T>(RouteData routeData, MenuNode<T> menuNode, Action<MenuNode<T>> failureAction) where T : struct
        {
            var node = menuNode.FindNode(routeData);
            if (node == null)
            {
                return null;
            }

            if (node.IsActive == false)
            {
                node.SetActive();
            }

            if (node.IsPageNode)
            {
                if (node.IsPermission == false)
                {
                    failureAction(node);
                }
                return node;
            }
            else if (node.IsAllow(node.ActionEnum) == false)
            {
                failureAction(node);
            }
            return node;
        }
    }
}
