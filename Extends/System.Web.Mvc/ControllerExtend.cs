using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace System.Web.Mvc
{
    /// <summary>
    /// 控制器扩展
    /// </summary>
    public static partial class ControllerExtend
    {
        /// <summary>
        /// 获取View的html
        /// </summary>
        /// <param name="controller">控制器</param>
        /// <returns></returns>
        public static string GetViewHtml(this Controller controller)
        {
            return controller.GetViewHtml(null, null);
        }

        /// <summary>
        /// 获取View的html
        /// </summary>
        /// <param name="controller">控制器</param>
        /// <param name="viewName">view名称</param>
        /// <returns></returns>
        public static string GetViewHtml(this Controller controller, string viewName)
        {
            return controller.GetViewHtml(viewName, null);
        }

        /// <summary>
        /// 获取View的html
        /// </summary>
        /// <param name="controller">控制器</param>
        /// <param name="model">实体</param>
        /// <returns></returns>
        public static string GetViewHtml(this Controller controller, object model)
        {
            return controller.GetViewHtml(null, model);
        }

        /// <summary>
        /// 获取View的html
        /// </summary>
        /// <param name="controller">控制器</param>
        /// <param name="viewName">view名称</param>
        /// <param name="model">实体</param>
        /// <returns></returns>
        public static string GetViewHtml(this Controller controller, string viewName, object model)
        {
            if (string.IsNullOrEmpty(viewName))
            {
                viewName = controller.ControllerContext.RouteData.GetRequiredString("action");
            }

            using (var sw = new StringWriter())
            {
                controller.ViewData.Model = model;
                ViewEngineResult viewResult = ViewEngines.Engines.FindView(controller.ControllerContext, viewName, null);
                ViewContext viewContext = new ViewContext(controller.ControllerContext, viewResult.View, controller.ViewData, controller.TempData, sw);
                viewResult.View.Render(viewContext, sw);
                return sw.GetStringBuilder().ToString();
            }
        }
    }
}
