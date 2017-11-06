using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Routing;

namespace Infrastructure.Menu
{
    /// <summary>
    /// 表示一个菜单节点
    /// 不可继承
    /// </summary>
    [Serializable]
    public sealed class MenuNode<T> where T : struct
    {
        /// <summary>
        /// 获取节点名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取节点图标
        /// </summary>
        public string Icon { get; set; }


        /// <summary>
        /// 获取是否是许可的
        /// </summary>
        [System.Xml.Serialization.XmlIgnore]
        public bool IsPermission { get; set; }

        /// <summary>
        /// 获取是否活动状态
        /// </summary>
        [Newtonsoft.Json.JsonIgnore]
        [System.Xml.Serialization.XmlIgnore]
        public bool IsActive { get; set; }


        /// <summary>
        /// 获取是否页面节点
        /// </summary>
        public bool IsPageNode { get; set; }


        /// <summary>
        /// 获取控制器名
        /// </summary>
        public string Controller { get; set; }

        /// <summary>
        /// 获取Action名
        /// </summary>
        public string ActionName { get; set; }

        /// <summary>
        /// 获取配置的操作行为枚举
        /// </summary>
        public T ActionEnum { get; set; }

        /// <summary>
        /// 获取允许的操作行为枚举
        /// </summary>
        [System.Xml.Serialization.XmlIgnore]
        public T AllowedAction { get; set; }

        /// <summary>
        /// 获取节点的父节点
        /// </summary>
        [Newtonsoft.Json.JsonIgnore]
        [System.Xml.Serialization.XmlIgnore]
        public MenuNode<T> ParentNode { get; set; }

        /// <summary>
        /// 获取节点的子节点
        /// </summary>
        public List<MenuNode<T>> ChildNodes { get; set; }

        /// <summary>
        /// 获取或设置附加数据
        /// </summary>
        public object Tag { get; set; }


        /// <summary>
        /// 获取子节点
        /// </summary>
        /// <param name="name">子节点名称</param>
        /// <returns></returns>
        public MenuNode<T> this[string name]
        {
            get
            {
                return this.ChildNodes.FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
            }
        }

        /// <summary>
        /// 菜单节点
        /// </summary>
        public MenuNode()
        {
            this.ChildNodes = new List<MenuNode<T>>();
        }

        /// <summary>
        /// 添加子节点
        /// </summary>
        /// <param name="nodes">节点</param>
        public MenuNode<T> AddChildNode(MenuNode<T> node)
        {
            node.ParentNode = this;

            if (this.IsPageNode == true)
            {
                var actionEnum = this.ActionEnum.GetHashCode() + node.ActionEnum.GetHashCode();
                this.ActionEnum = (T)Enum.Parse(typeof(T), actionEnum.ToString());
            }

            this.ChildNodes.Add(node);
            return node;
        }

        /// <summary>
        /// 从父层节点的子节点中移除
        /// </summary>
        public void Remove()
        {
            if (this.ParentNode != null)
            {
                this.ParentNode.ChildNodes.Remove(this);
            }
        }

        /// <summary>
        /// 获取节点包含的所有Page节点
        /// </summary>
        /// <returns></returns>
        public IList<MenuNode<T>> GetPageNodes()
        {
            var list = new List<MenuNode<T>>();
            this.AddNodeToList(this, list);
            return list;
        }

        /// <summary>
        /// 将节点添加到列表中
        /// </summary>
        /// <param name="node"></param>
        /// <param name="list"></param>
        private void AddNodeToList(MenuNode<T> node, List<MenuNode<T>> list)
        {
            if (node.IsPageNode)
            {
                list.Add(node);
            }

            foreach (var item in node.ChildNodes)
            {
                this.AddNodeToList(item, list);
            }
        }

        /// <summary>
        /// 设置node和所有父级
        /// </summary>
        /// <param name="node"></param>
        /// <param name="action"></param>
        private void SetWithParents(MenuNode<T> node, Action<MenuNode<T>> action)
        {
            while (node != null)
            {
                action.Invoke(node);
                node = node.ParentNode;
            }
        }

        /// <summary>
        /// 调整授权信息      
        /// </summary>      
        /// <param name="value">是否授权</param>
        /// <returns></returns>
        public MenuNode<T> SetPermission(bool value)
        {
            var pageNodes = this.GetPageNodes();
            foreach (var node in pageNodes)
            {
                node.AllowedAction = value ? node.ActionEnum : default(T);
                this.SetWithParents(node, (n) => n.IsPermission = value);
            }
            return this;
        }

        /// <summary>
        /// 调整授权信息
        /// 包括配置的行为和允许的行为
        /// </summary>      
        /// <param name="permissionNodes">许可的节点信息</param>
        /// <returns></returns>
        public MenuNode<T> SetPermission(IEnumerable<PermissionNode<T>> permissionNodes)
        {
            if (permissionNodes == null)
            {
                return this.SetPermission(false);
            }

            var pageNodes = this.GetPageNodes();
            foreach (var node in pageNodes)
            {
                var pNode = permissionNodes.FirstOrDefault(i => i.HashMd5 == node.GetHashMd5());
                node.IsPermission = pNode != null;
                node.AllowedAction = pNode == null ? default(T) : pNode.AllowedAction;

                this.SetWithParents(node.ParentNode, (p) => p.IsPermission = p.ChildNodes.Any(c => c.IsPermission));
            }

            return this;
        }


        /// <summary>
        /// 设置为激活状态
        /// 节点的所有父节点也跟随激活
        /// </summary>
        /// <returns></returns>
        public MenuNode<T> SetActive()
        {
            this.SetWithParents(this, (p) => p.IsActive = true);
            return this;
        }

        /// <summary>
        /// 获取节点是否允许某种操作行为
        /// </summary>
        /// <param name="action">行为</param>
        /// <returns></returns>
        public bool IsAllow(T action)
        {
            var actionEnum = action.GetHashCode();
            if (actionEnum == 0)
            {
                return true;
            }

            var allowedAction = 0;
            if (this.IsPermission == true)
            {
                allowedAction = this.AllowedAction.GetHashCode();
            }
            else
            {
                allowedAction = this.ParentNode.AllowedAction.GetHashCode();
            }
            return (allowedAction & actionEnum) > 0;
        }


        /// <summary>
        /// 查找页面节点
        /// </summary>
        /// <param name="name">节点名称L</param>        
        /// <returns></returns>
        public MenuNode<T> FindPageNode(string name)
        {
            return this.GetPageNodes().FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 查找页面节点
        /// </summary>
        /// <param name="controller">控制器</param>
        /// <param name="action">Action名</param>
        /// <returns></returns>
        public MenuNode<T> FindPageNode(string controller, string action)
        {
            Func<MenuNode<T>, bool> predicate = item =>
                string.Equals(item.Controller, controller, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(item.ActionName, action, StringComparison.OrdinalIgnoreCase);

            return this.GetPageNodes().FirstOrDefault(predicate);
        }

        /// <summary>
        /// 查找页面节点
        /// </summary>
        /// <param name="routeData">路由数据</param>      
        /// <returns></returns>
        public MenuNode<T> FindPageNode(RouteData routeData)
        {
            var controller = routeData.Values["controller"].ToString();
            var actionName = routeData.Values["action"].ToString();

            return this.FindPageNode(controller, actionName);
        }

        /// <summary>
        /// 查找页面节点或行为节点
        /// </summary>
        /// <param name="routeData">路由数据</param>
        /// <returns></returns>
        public MenuNode<T> FindNode(RouteData routeData)
        {
            var controller = routeData.Values["controller"].ToString();
            var actionName = routeData.Values["action"].ToString();

            Func<MenuNode<T>, bool> predicate = item =>
              string.Equals(item.Controller, controller, StringComparison.OrdinalIgnoreCase) &&
              string.Equals(item.ActionName, actionName, StringComparison.OrdinalIgnoreCase);

            var pageNodes = this.GetPageNodes();
            var searchNodes = pageNodes.Concat(pageNodes.SelectMany(p => p.ChildNodes));

            return searchNodes.FirstOrDefault(predicate);
        }


        /// <summary>
        /// 过滤节点
        /// </summary>
        /// <param name="nodeMd5">保留的菜单节点的md5</param>
        /// <returns></returns>
        public MenuNode<T> Filter(IEnumerable<FilterNode<T>> filterNodes)
        {
            var parentNodes = this.GetPageNodes().Select(item => item.ParentNode).Distinct().ToArray();
            foreach (var parentNode in parentNodes)
            {
                var unFilters = this.UpdatePageNodes(parentNode, filterNodes).ToArray();
                foreach (var item in unFilters)
                {
                    parentNode.ChildNodes.Remove(item);
                }

                this.SetWithParents(parentNode, (node) =>
                {
                    if (node.ChildNodes.Count == 0)
                    {
                        node.Remove();
                    }
                });
            }
            return this;
        }

        /// <summary>
        /// 过滤page节点
        /// </summary>
        /// <param name="parentNode"></param>
        /// <param name="filterNodes"></param>
        /// <returns>返回没更新的节点</returns>
        private IEnumerable<MenuNode<T>> UpdatePageNodes(MenuNode<T> parentNode, IEnumerable<FilterNode<T>> filterNodes)
        {
            foreach (var node in parentNode.ChildNodes)
            {
                var md5 = node.GetHashMd5();
                var filterNode = filterNodes.FirstOrDefault(item => item.HashMd5 == md5);
                if (filterNode != null)
                {
                    node.ActionEnum = filterNode.ActionEnum;
                }
                else
                {
                    yield return node;
                }
            }
        }

        /// <summary>
        /// 转换为PermissionNode对象
        /// </summary>
        /// <returns></returns>
        public PermissionNode<T> ToPermissionNode()
        {
            return new PermissionNode<T> { HashMd5 = this.GetHashMd5(), AllowedAction = this.AllowedAction };
        }

        /// <summary>
        /// 获取完整名称
        /// </summary>
        /// <returns></returns>
        public string GetFullName()
        {
            return string.Join(".", this.GetNames().Reverse());
        }

        /// <summary>
        /// 获取名称
        /// </summary>
        /// <returns></returns>
        private IEnumerable<string> GetNames()
        {
            var parent = this;
            while (parent != null)
            {
                yield return parent.Name;
                parent = parent.ParentNode;
            }
        }

        /// <summary>
        /// 获取节点的MD5值
        /// </summary>
        /// <returns></returns>
        public string GetHashMd5()
        {
            var code = this.GetFullName() + this.Controller + this.ActionName;
            return Utility.Encryption.GetMD5(code.ToLower());
        }

        /// <summary>
        /// 获取哈希码
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return this.GetHashMd5().GetHashCode();
        }

        /// <summary>
        /// 比较是否和目标相等
        /// </summary>
        /// <param name="obj">目标对象</param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return obj != null && this.GetHashCode() == obj.GetHashCode();
        }

        /// <summary>
        /// 字符串显示
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            if (this.Name.IsNullOrEmpty())
            {
                return "ActionEnum：" + this.ActionEnum.ToString();
            }
            return this.Name;
        }
    }
}
