using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;

namespace Infrastructure.Menu
{
    /// <summary>
    /// 提供菜单生成器
    /// </summary>
    public class MenuBuilder<T> where T : struct
    {
        /// <summary>
        /// 顶层节点
        /// </summary>
        private MenuNode<T> topNode = new MenuNode<T>();

        /// <summary>
        /// 获取最后System节点
        /// </summary>
        private MenuNode<T> LastSystem
        {
            get
            {
                return this.topNode.ChildNodes.LastOrDefault();
            }
        }

        /// <summary>
        /// 获取最后Module节点
        /// </summary>
        private MenuNode<T> LastModule
        {
            get
            {
                return this.LastSystem.ChildNodes.LastOrDefault();
            }
        }

        /// <summary>
        /// 获取最后Page节点
        /// </summary>
        private MenuNode<T> LastPage
        {
            get
            {
                return this.LastModule.ChildNodes.LastOrDefault();
            }
        }

        /// <summary>
        /// 菜单生成器
        /// </summary>
        public MenuBuilder()
        {
            if (typeof(T).IsEnum == false)
            {
                throw new Exception("泛型参数类型必须为枚举类型");
            }
        }

        /// <summary>
        /// 添加子system节点
        /// </summary>
        /// <param name="systemNode">system节点</param>
        /// <returns></returns>
        public MenuBuilder<T> AddSystem(MenuNode<T> systemNode)
        {
            this.topNode.AddChildNode(systemNode);
            return this;
        }

        /// <summary>
        /// 添加module节点
        /// </summary>
        /// <param name="moduleNode">module节点</param>
        /// <returns></returns>
        public MenuBuilder<T> AddModule(MenuNode<T> moduleNode)
        {
            this.LastSystem.AddChildNode(moduleNode);
            return this;
        }

        /// <summary>
        /// 添加菜单节点
        /// </summary>
        /// <param name="pageNode">page节点</param>
        /// <returns></returns>
        public MenuBuilder<T> AddPage(MenuNode<T> pageNode)
        {
            pageNode.Controller = pageNode.Controller.ToLowerIfNoNull();
            pageNode.ActionName = pageNode.ActionName.ToLowerIfNoNull();
            pageNode.IsPageNode = true;
            this.LastModule.AddChildNode(pageNode);
            return this;
        }


        /// <summary>
        /// 添加action节点
        /// </summary>
        /// <param name="actioNode">action节点</param>
        /// <returns></returns>
        public MenuBuilder<T> AddAction(MenuNode<T> actioNode)
        {
            actioNode.Controller = actioNode.Controller.ToLowerIfNoNull();
            actioNode.ActionName = actioNode.ActionName.ToLowerIfNoNull();
            this.LastPage.AddChildNode(actioNode);
            return this;
        }


        /// <summary>
        /// 添加子系统
        /// </summary>
        /// <param name="name">系统名</param>
        /// <returns></returns>
        public MenuBuilder<T> System(string name)
        {
            var system = new MenuNode<T> { Name = name };
            return this.AddSystem(system);
        }

        /// <summary>
        /// 添加子系统
        /// </summary>
        /// <param name="name">系统名</param>
        /// <param name="icon">图标</param>
        /// <returns></returns>
        public MenuBuilder<T> System(string name, string icon)
        {
            var system = new MenuNode<T> { Name = name, Icon = icon };
            return this.AddSystem(system);
        }


        /// <summary>
        /// 添加子系统
        /// </summary>
        /// <param name="name">系统名</param>
        /// <param name="tag">附加数据</param>
        /// <returns></returns>
        public MenuBuilder<T> System(string name, object tag)
        {
            var system = new MenuNode<T> { Name = name, Tag = tag };
            return this.AddSystem(system);
        }

        /// <summary>
        /// 添加子系统
        /// </summary>
        /// <param name="name">系统名</param>
        /// <param name="icon">图标</param>
        /// <param name="tag">附加数据</param>
        /// <returns></returns>
        public MenuBuilder<T> System(string name, string icon, object tag)
        {
            var system = new MenuNode<T> { Name = name, Icon = icon, Tag = tag };
            return this.AddSystem(system);
        }

        /// <summary>
        /// 添加模块节点
        /// </summary>
        /// <param name="name">模块名</param>
        /// <returns></returns>
        public MenuBuilder<T> Module(string name)
        {
            var module = new MenuNode<T> { Name = name };
            return this.AddModule(module);
        }

        /// <summary>
        /// 添加模块节点
        /// </summary>
        /// <param name="name">模块名</param>
        /// <param name="icon">图标</param>
        /// <returns></returns>
        public MenuBuilder<T> Module(string name, string icon)
        {
            var module = new MenuNode<T> { Name = name, Icon = icon };
            return this.AddModule(module);
        }


        /// <summary>
        /// 添加菜单节点
        /// </summary>
        /// <param name="name">菜单名称</param>
        /// <param name="controller">控制器</param>
        /// <param name="actionName">Action名</param>
        /// <returns></returns>
        public MenuBuilder<T> Page(string name, string controller, string actionName)
        {
            var menu = new MenuNode<T> { Name = name, Controller = controller, ActionName = actionName };
            return this.AddPage(menu);
        }

        /// <summary>
        /// 添加菜单节点
        /// </summary>
        /// <param name="name">菜单名称</param>
        /// <param name="controller">控制器</param>
        /// <param name="actionName">Action名</param>
        /// <param name="icon">图标</param>
        /// <returns></returns>
        public MenuBuilder<T> Page(string name, string controller, string actionName, string icon)
        {
            var menu = new MenuNode<T> { Name = name, Controller = controller, ActionName = actionName, Icon = icon };
            return this.AddPage(menu);
        }

        /// <summary>
        /// 添加行为节点
        /// </summary>    
        /// <param name="controller">控制器</param>
        /// <param name="actionName">Action名</param>
        /// <returns></returns>
        public MenuBuilder<T> Action(string controller, string actionName)
        {
            var mainMenu = new MenuNode<T> { Controller = controller, ActionName = actionName };
            return this.AddAction(mainMenu);
        }


        /// <summary>
        /// 添加行为节点
        /// </summary>    
        /// <param name="controller">控制器</param>
        /// <param name="actionName">Action名</param>
        /// <param name="actionEnum">行为枚举</param>
        /// <returns></returns>
        public MenuBuilder<T> Action(string controller, string actionName, T actionEnum)
        {
            var mainMenu = new MenuNode<T> { Controller = controller, ActionName = actionName, ActionEnum = actionEnum };
            return this.AddAction(mainMenu);
        }

        /// <summary>
        /// 获取顶层节点
        /// </summary>
        /// <returns></returns>
        public MenuNode<T> GetTopNode()
        {
            return this.topNode;
        }
    }
}
