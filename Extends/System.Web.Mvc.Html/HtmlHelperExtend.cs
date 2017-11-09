using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Optimization;


namespace System.Web.Mvc.Html
{
    /// <summary>
    /// HTML扩展
    /// </summary>
    public static partial class HtmlHelperExtend
    {
        /// <summary>
        /// 获取表达式的属性的值和属性名
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <typeparam name="TKey">属性类型</typeparam>
        /// <param name="helper">htmlHelper</param>
        /// <param name="expression">表达式</param>
        /// <param name="name">属性名</param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        /// <returns></returns>
        private static TKey GetExpressionValue<T, TKey>(this HtmlHelper<T> helper, Expression<Func<T, TKey>> expression, out string name)
        {
            if (expression == null)
            {
                throw new ArgumentNullException();
            }

            MemberExpression body = expression.Body as MemberExpression;
            if (body == null || body.Member.DeclaringType.IsAssignableFrom(typeof(T)) == false || body.Expression.NodeType != ExpressionType.Parameter)
            {
                throw new ArgumentException();
            }

            name = body.Member.Name;
            var value = typeof(T).GetProperty(name).GetValue(helper.ViewData.Model, null);
            return (TKey)value;
        }

        /// <summary>
        /// bool类型生成Select下拉框
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="helper">html生成类</param>
        /// <param name="expression">选择器</param>
        /// <param name="trueText">为true时的文本</param>
        /// <param name="falseText">为false时的文本</param>
        /// <param name="htmlAttributes">html属性</param>       
        /// <returns></returns>
        public static MvcHtmlString DropDownListFor<T>(this HtmlHelper<T> helper, Expression<Func<T, bool>> expression, string trueText, string falseText, object htmlAttributes = null)
        {
            string name = string.Empty;
            var value = helper.GetExpressionValue<T, bool>(expression, out name);

            var list = new List<SelectListItem>()
            {
                new SelectListItem() { Value = "true", Text = trueText, Selected = value == true },
                new SelectListItem() { Value = "false", Text = falseText, Selected = value == false }
            };
            return helper.DropDownList(name, list, htmlAttributes);
        }


        /// <summary>
        /// bool类型生成Select下拉框
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="helper">html生成类</param>
        /// <param name="expression">选择器</param>
        /// <param name="trueText">为true时的文本</param>
        /// <param name="falseText">为false时的文本</param>
        /// <param name="htmlAttributes">html属性</param>       
        /// <returns></returns>
        public static MvcHtmlString DropDownListFor<T>(this HtmlHelper<T> helper, Expression<Func<T, bool>> expression, string trueText, string falseText, IDictionary<string, object> htmlAttributes)
        {
            string name = string.Empty;
            var value = helper.GetExpressionValue<T, bool>(expression, out name);

            var list = new List<SelectListItem>()
            {
                new SelectListItem() { Value = "true", Text = trueText, Selected = value == true },
                new SelectListItem() { Value = "false", Text = falseText, Selected = value == false }
            };
            return helper.DropDownList(name, list, htmlAttributes);
        }

        /// <summary>
        /// 生成Select下拉框
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="helper"></param>
        /// <param name="expression">表达式</param>
        /// <param name="selectListItems">选项项</param>
        /// <param name="optionLabel">可行标签</param>
        /// <param name="htmlAttributes">html属性</param>
        /// <returns></returns>
        public static MvcHtmlString DropDownListFor<T, TKey>(this HtmlHelper<T> helper, Expression<Func<T, TKey>> expression, IEnumerable<KeyValuePair<string, string>> selectListItems, string optionLabel, object htmlAttributes)
        {
            var selectList = selectListItems.Select(item => new SelectListItem { Text = item.Value, Value = item.Key });
            return helper.DropDownListFor(expression, selectList, optionLabel, htmlAttributes);
        }

        /// <summary>
        /// 生成Select下拉框
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="helper"></param>
        /// <param name="expression">表达式</param>
        /// <param name="selectListItems">选项项</param>
        /// <param name="optionLabel">可行标签</param>
        /// <param name="htmlAttributes">html属性</param>
        /// <returns></returns>
        public static MvcHtmlString DropDownListFor<T, TKey>(this HtmlHelper<T> helper, Expression<Func<T, TKey>> expression, IEnumerable<KeyValuePair<string, string>> selectListItems, string optionLabel, IDictionary<string, object> htmlAttributes)
        {
            var selectList = selectListItems.Select(item => new SelectListItem { Text = item.Value, Value = item.Key });
            return helper.DropDownListFor(expression, selectList, optionLabel, htmlAttributes);
        }

        /// <summary>
        /// 生成Select下拉框
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="helper"></param>
        /// <param name="expression">表达式</param>
        /// <param name="selectListItems">选项项</param>        
        /// <param name="htmlAttributes">html属性</param>
        /// <returns></returns>
        public static MvcHtmlString DropDownListFor<T, TKey>(this HtmlHelper<T> helper, Expression<Func<T, TKey>> expression, IEnumerable<KeyValuePair<string, string>> selectListItems, IDictionary<string, object> htmlAttributes)
        {
            var selectList = selectListItems.Select(item => new SelectListItem { Text = item.Value, Value = item.Key });
            return helper.DropDownListFor(expression, selectList, htmlAttributes);
        }

        /// <summary>
        /// 生成Select下拉框
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="helper"></param>
        /// <param name="expression">表达式</param>
        /// <param name="selectListItems">选项项</param>       
        /// <param name="htmlAttributes">html属性</param>
        /// <returns></returns>
        public static MvcHtmlString DropDownListFor<T, TKey>(this HtmlHelper<T> helper, Expression<Func<T, TKey>> expression, IEnumerable<KeyValuePair<string, string>> selectListItems, object htmlAttributes)
        {
            var selectList = selectListItems.Select(item => new SelectListItem { Text = item.Value, Value = item.Key });
            return helper.DropDownListFor(expression, selectList, htmlAttributes);
        }

        /// <summary>
        /// 生成Select下拉框
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="helper"></param>
        /// <param name="expression">表达式</param>
        /// <param name="selectListItems">选项项</param>  
        /// <returns></returns>
        public static MvcHtmlString DropDownListFor<T, TKey>(this HtmlHelper<T> helper, Expression<Func<T, TKey>> expression, IEnumerable<KeyValuePair<string, string>> selectListItems)
        {
            var selectList = selectListItems.Select(item => new SelectListItem { Text = item.Value, Value = item.Key });
            return helper.DropDownListFor(expression, selectList);
        }

        /// <summary>
        /// 将html属性合并到字典中
        /// </summary>
        /// <param name="dic">字典</param>
        /// <param name="attribute">属性</param>
        private static void MergenAttribute(IDictionary<string, object> dic, object attribute)
        {
            if (attribute == null)
            {
                return;
            }

            var properties = attribute.GetType().GetProperties();
            foreach (var p in properties)
            {
                var key = p.Name.Replace("_", "-").ToLower();
                var value = p.GetValue(attribute, null);

                if (dic.ContainsKey(key))
                {
                    if (key == "class")
                    {
                        dic[key] = string.Format("{0} {1}", dic[key], value).Trim();
                    }
                }
                else
                {
                    dic.Add(key, value);
                }
            }
        }


        /// <summary>
        /// 只读文本框
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TProperty"></typeparam>
        /// <param name="htmlHelper"></param>
        /// <param name="expression"></param>
        /// <param name="htmlAttributes"></param>
        /// <returns></returns>
        public static MvcHtmlString TextBoxReadonlyFor<TModel, TProperty>(this HtmlHelper<TModel> htmlHelper, Expression<Func<TModel, TProperty>> expression, object htmlAttributes)
        {
            var dic = new Dictionary<string, object>();
            MergenAttribute(dic, htmlAttributes);
            MergenAttribute(dic, new { @readonly = "readonly" });
            return htmlHelper.TextBoxFor(expression, dic);
        }

        /// <summary>
        /// 只读文本框
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TProperty"></typeparam>
        /// <param name="htmlHelper"></param>
        /// <param name="expression"></param>
        /// <param name="htmlAttributes"></param>
        /// <returns></returns>
        public static MvcHtmlString TextBoxReadonlyFor<TModel, TProperty>(this HtmlHelper<TModel> htmlHelper, Expression<Func<TModel, TProperty>> expression, IDictionary<string, object> htmlAttributes)
        {
            var dic = htmlAttributes == null ? new Dictionary<string, object>() : htmlAttributes;
            MergenAttribute(dic, new { @readonly = "readonly" });
            return htmlHelper.TextBoxFor(expression, dic);
        }
        /// <summary>
        /// 只读文本框
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <typeparam name="TProperty"></typeparam>
        /// <param name="htmlHelper"></param>
        /// <param name="expression"></param>
        /// <param name="htmlAttributes"></param>
        /// <returns></returns>
        public static MvcHtmlString TextBoxReadonlyFor<TModel, TProperty>(this HtmlHelper<TModel> htmlHelper, Expression<Func<TModel, TProperty>> expression)
        {
            return htmlHelper.TextBoxFor(expression, new { @readonly = "readonly" });
        }

        /// <summary>
        /// 动态捆绑多个脚本
        /// </summary>
        /// <param name="htmlHelper"></param>
        /// <param name="scripts">javscript</param>
        /// <returns></returns>
        public static IHtmlString BundleScripts(this HtmlHelper htmlHelper, params Func<object, object>[] scripts)
        {
            if (scripts == null)
            {
                throw new ArgumentNullException("scripts");
            }

            var inputs = new StringBuilder();
            foreach (var script in scripts)
            {
                inputs.AppendLine(script.Invoke(null).ToString().ToLower());
            }

            var applicationPath = htmlHelper.ViewContext.HttpContext.Request.ApplicationPath.ToLower();
            Func<string, string> fixSrc = (src) => applicationPath == "/" ? "~" + src : src.Replace(applicationPath, "~");

            var srcs = inputs.ToString().Matches(@"(?<=src="").+?\.js(?="")").Select(item => fixSrc(item)).ToArray();
            var path = string.Format("~/{0}.js", Math.Abs(string.Join(string.Empty, srcs).GetHashCode()));

            if (BundleTable.Bundles.GetBundleFor(path) == null)
            {
                BundleTable.Bundles.Add(new ScriptBundle(path).Include(srcs));
            }
            return Scripts.Render(path);
        }

        /// <summary>
        /// 动态捆绑多个样式
        /// </summary>
        /// <param name="htmlHelper"></param>
        /// <param name="styles">样式</param>
        /// <returns></returns>
        public static IHtmlString BundleStyles(this HtmlHelper htmlHelper, params Func<object, object>[] styles)
        {
            if (styles == null)
            {
                throw new ArgumentNullException("styles");
            }

            var inputs = new StringBuilder();
            foreach (var style in styles)
            {
                inputs.AppendLine(style.Invoke(null).ToString().ToLower());
            }

            var applicationPath = htmlHelper.ViewContext.HttpContext.Request.ApplicationPath.ToLower();
            Func<string, string> fixHref = (href) => applicationPath == "/" ? "~" + href : href.Replace(applicationPath, "~");

            var hrefs = inputs.ToString().Matches(@"(?<=href="").+?\.css(?="")").Select(item => fixHref(item)).ToArray();
            var path = string.Format("~/{0}.css", Math.Abs(string.Join(string.Empty, hrefs).GetHashCode()));

            if (BundleTable.Bundles.GetBundleFor(path) == null)
            {
                Bundle bundle = new StyleBundle(path);
                foreach (var href in hrefs)
                {
                    bundle = bundle.Include(href, new StyleUrlTransform(applicationPath));
                }
                BundleTable.Bundles.Add(bundle);
            }
            return Styles.Render(path);
        }
    }
}