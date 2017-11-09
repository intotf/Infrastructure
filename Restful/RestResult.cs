using Infrastructure.Page;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace System.Web
{
    /// <summary>
    /// 表示Restful风格结果
    /// </summary>
    public class RestResult : ActionResult
    {
        /// <summary>
        /// 模型
        /// </summary>
        private RestModel model;

        /// <summary>
        /// Restful风格结果
        /// </summary>
        /// <param name="model">模型</param>
        public RestResult(RestModel model)
        {
            this.model = model;
        }

        /// <summary>
        /// Restful风格结果
        /// </summary>
        /// <param name="state">状态</param>
        /// <param name="msg">消息</param>
        public RestResult(bool state, string msg)
        {
            this.model = new RestModel { State = state, Msg = msg };
        }

        /// <summary>
        /// 执行结果
        /// </summary>
        /// <param name="context">内容上下文</param>
        public override void ExecuteResult(ControllerContext context)
        {
            var contentType = "application/json";
            var json = Infrastructure.Utility.JsonSerializer.Serialize(this.model);
            var callback = context.RequestContext.HttpContext.Request["callback"];
            if (callback == null)
            {
                var result = new ContentResult() { ContentType = contentType, Content = json };
                result.ExecuteResult(context);
            }
            else
            {
                var result = new ContentResult { Content = string.Format("{0}({1})", callback, json) };
                result.ExecuteResult(context);
            }
        }

        /// <summary>
        /// 表示成功的结果
        /// </summary>
        /// <param name="data">data字段</param>
        /// <returns></returns>
        public static RestResult True(object data)
        {
            var model = new RestModel { State = true, Data = data };
            return new RestResult(model);
        }

        /// <summary>
        /// 表示成功的结果
        /// </summary>
        /// <param name="data">data字段</param>
        /// <param name="msg">消息</param>
        /// <returns></returns>
        public static RestResult True(object data, string msg)
        {
            var model = new RestModel { State = true, Data = data, Msg = msg };
            return new RestResult(model);
        }

        /// <summary>
        /// 表示失败的结果
        /// </summary>
        /// <param name="msg">消息</param>
        /// <returns></returns>
        public static RestResult False(string msg)
        {
            var model = new RestModel { Msg = msg };
            return new RestResult(model);
        }

        /// <summary>
        /// 表示失败的结果
        /// </summary>
        /// <param name="msg">消息</param>
        /// <param name="data">data字段</param>
        /// <returns></returns>
        public static RestResult False(string msg, object data)
        {
            var model = new RestModel { Msg = msg, Data = data };
            return new RestResult(model);
        }

        /// <summary>
        /// 表示失败的结果
        /// </summary>
        /// <param name="code">错误码</param>
        /// <returns></returns>
        public static RestResult False(ErrorCode code)
        {
            var model = new RestModel { Msg = code.GetFieldDisplay(), Code = code.GetHashCode() };
            return new RestResult(model);
        }

        /// <summary>
        /// 从page转换得到
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="page"></param>
        /// <returns></returns>
        public static RestResult FromPage<T>(IPageInfo<T> page) where T : class
        {
            var data = new { count = page.TotalCount, list = page.ToArray() };
            return RestResult.True(data);
        }
    }
}