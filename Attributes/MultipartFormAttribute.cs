using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Mvc;

namespace Infrastructure.Attributes
{
    /// <summary>
    /// 表示二进制数据表单
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    public class MultipartFormAttribute : ActionMethodSelectorAttribute
    {
        public override bool IsValidForRequest(ControllerContext controllerContext, System.Reflection.MethodInfo methodInfo)
        {
            var request = controllerContext.HttpContext.Request;
            if (string.Equals(request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase) == false)
            {
                return true;
            }

            var contentType = request.Headers["Content-Type"];
            if (string.IsNullOrEmpty(contentType) == false)
            {
                return Regex.IsMatch(contentType, "multipart/form-data", RegexOptions.IgnoreCase);
            }
            return false;
        }
    }
}
