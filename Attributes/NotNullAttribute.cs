using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infrastructure.Attributes
{
    /// <summary>
    /// 表示参数不能为空的特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter)]
    public class NotNullAttribute : Attribute
    {
        /// <summary>
        /// 获取或设置错误提示信息
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}
