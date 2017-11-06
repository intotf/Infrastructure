using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Routing;

namespace Infrastructure.Menu
{
    /// <summary>
    /// 菜单元素集合
    /// </summary>
    public sealed class MenuNodeCollection<T> where T : struct
    {
        /// <summary>
        /// 用于保存节点数据
        /// </summary>
        private IEnumerable<MenuNode<T>> nodes;

        /// <summary>
        /// 菜单元素集合
        /// </summary>
        /// <param name="nodes"></param>
        internal MenuNodeCollection(IEnumerable<MenuNode<T>> nodes)
        {
            this.nodes = nodes;
        }


        /// <summary>
        /// 调整授权信息      
        /// </summary>      
        /// <param name="value">是否授权</param>
        /// <returns></returns>
        public MenuNodeCollection<T> SetPermission(bool value)
        {
            foreach (var node in this.nodes)
            {
                node.IsPermission = value;
                if (node.IsPermission == true)
                {
                    node.AllowedAction = node.ActionEnum;
                }
                else
                {
                    node.AllowedAction = default(T);
                }
            }

            return this.FixModeulePermission();
        }

        /// <summary>
        /// 调整授权信息
        /// 包括配置的行为和允许的行为
        /// </summary>      
        /// <param name="permissionNodes">许可的节点信息</param>
        /// <returns></returns>
        public MenuNodeCollection<T> SetPermission(IEnumerable<PermissionNode<T>> permissionNodes)
        {
            if (permissionNodes == null)
            {
                return this.SetPermission(false);
            }

            foreach (var node in this.nodes)
            {
                if (node.Level == NodeLevels.Menu)
                {
                    var pNode = permissionNodes.FirstOrDefault(i => i.HashMd5 == node.GetHashMd5());
                    node.IsPermission = pNode != null;
                    node.AllowedAction = pNode == null ? default(T) : pNode.AllowedAction;
                }
            }
            return this.FixModeulePermission();
        }


        /// <summary>
        /// 调模块的整授权信息       
        /// </summary>              
        /// <returns></returns>
        private MenuNodeCollection<T> FixModeulePermission()
        {
            foreach (var node in this.nodes)
            {
                if (node.Level == NodeLevels.Module)
                {
                    var firstChildNode = node.ChildNodes.FirstOrDefault(i => i.IsPermission);
                    node.IsPermission = firstChildNode != null;

                    if (node.IsPermission == true)
                    {
                        node.Controller = firstChildNode.Controller;
                        node.ActionName = firstChildNode.ActionName;
                    }
                }
            }

            return this;
        }


        /// <summary>
        /// 获取所有模块节点
        /// </summary>
        public IEnumerable<MenuNode<T>> GetModuleNodes()
        {
            return this.nodes.Where(i => i.Level == NodeLevels.Module);
        }

        /// <summary>
        /// 获取所有菜单节点
        /// </summary>
        /// <returns></returns>
        public IEnumerable<MenuNode<T>> GetMenuNodes()
        {
            return this.nodes.Where(i => i.Level == NodeLevels.Menu);
        }

        /// <summary>
        /// 查找节点
        /// 不包含模块节点
        /// </summary>
        /// <param name="predicate">条件</param>       
        /// <returns></returns>
        public MenuNode<T> FindNode(Func<MenuNode<T>, bool> predicate)
        {
            return this.nodes.Where(item => item.Level != NodeLevels.Module).Where(predicate).FirstOrDefault();
        }

        /// <summary>
        /// 查找节点
        /// 不包含模块节点
        /// </summary>
        /// <param name="name">节点名称或URL</param>        
        /// <returns></returns>
        public MenuNode<T> FindNode(string name)
        {
            return this.FindNode(item => item.Name == name || item.Controller + "/" + item.ActionName == name);
        }

        /// <summary>
        /// 查找节点
        /// 不包含模块节点
        /// </summary>
        /// <param name="controller">控制器</param>
        /// <param name="action">Action名</param>
        /// <returns></returns>
        public MenuNode<T> FindNode(string controller, string action)
        {
            controller = controller.ToLowerIfNoNull();
            action = action.ToLowerIfNoNull();
            Func<MenuNode<T>, bool> predicate = item => item.Controller == controller && item.ActionName == action;
            return this.FindNode(predicate);
        }


        /// <summary>
        /// 查找节点      
        /// 不包含模块节点
        /// </summary>
        /// <param name="routeData">路由数据</param>      
        /// <returns></returns>
        public MenuNode<T> FindNode(RouteData routeData)
        {
            var controller = routeData.Values["controller"].ToString();
            var actionName = routeData.Values["action"].ToString();
            return this.FindNode(controller, actionName);
        }

        /// <summary>
        /// 过滤节点
        /// </summary>
        /// <param name="nodeMd5">保留的节点的md5</param>
        /// <returns></returns>
        public MenuNodeCollection<T> Filter(IEnumerable<FilterNode<T>> filterNodes)
        {
            var menuNodes = this.GetModuleNodes().SelectMany(m => this.FilterModule(m, filterNodes).ToArray());
            return new MenuNodeCollection<T>(menuNodes.ToArray());
        }

        /// <summary>
        /// 过滤模块
        /// </summary>
        /// <param name="module">模块</param>
        /// <param name="filterNodes">保留的菜单节点</param>
        /// <returns></returns>
        private IEnumerable<MenuNode<T>> FilterModule(MenuNode<T> module, IEnumerable<FilterNode<T>> filterNodes)
        {
            var childNodes = this.FilterChildNodes(module, filterNodes).ToArray();
            if (childNodes.Length == 0)
            {
                yield break;
            }

            module.ChildNodes = childNodes;
            yield return module;

            foreach (var node in childNodes)
            {
                yield return node;
            }
        }

        /// <summary>
        /// 过滤模块的子节点
        /// </summary>
        /// <param name="module">模块</param>
        /// <param name="nodeMd5">保留的节点</param>
        /// <returns></returns>
        private IEnumerable<MenuNode<T>> FilterChildNodes(MenuNode<T> module, IEnumerable<FilterNode<T>> filterNodes)
        {
            foreach (var node in module.ChildNodes)
            {
                var filterNode = filterNodes.FirstOrDefault(item => item.HashMd5 == node.GetHashMd5());
                if (filterNode != null)
                {
                    node.ActionEnum = filterNode.ActionEnum;
                    yield return node;
                }
            }
        }

        /// <summary>
        /// 转换为可迭代类型
        /// </summary>
        /// <returns></returns>
        public IEnumerable<MenuNode<T>> AsEnumerable()
        {
            return this.nodes;
        }
    }
}
