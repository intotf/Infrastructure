using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace System.Web
{
    /// <summary>
    /// Rest模型
    /// </summary>
    public class RestModel
    {
        /// <summary>
        /// 状态
        /// </summary>
        public bool State { get; set; }

        /// <summary>
        /// 错误码 0无错误 1为Token无效 2为终端被禁用
        /// </summary>
        public int Code { get; set; }

        /// <summary>
        /// 提示信息
        /// </summary>
        public string Msg { get; set; }

        /// <summary>
        /// 业务类型
        /// </summary>
        public object Data { get; set; }
    }
}
