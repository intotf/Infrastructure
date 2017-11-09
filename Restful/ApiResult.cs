using Infrastructure.Utility;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace System
{
    /// <summary>
    /// 表示Rest风格Api结果
    /// </summary>
    [DebuggerDisplay("State = {State}")]
    public class ApiResult
    {
        /// <summary>
        /// 状态
        /// </summary>
        public bool State { get; set; }

        /// <summary>
        /// 状态码
        /// </summary>
        public ErrorCode Code { get; set; }

        /// <summary>
        /// 提示信息
        /// </summary>
        public string Msg { get; set; }

        /// <summary>
        /// 业务数据
        /// </summary>
        public dynamic Data { get; set; }

        /// <summary>
        /// 转换为对应的数据类型
        /// </summary>
        /// <typeparam name="TData">Data的类型</typeparam>
        /// <returns></returns>
        public ApiResult<TData> As<TData>()
        {
            var result = new ApiResult<TData>
            {
                State = this.State,
                Msg = this.Msg,
                Code = this.Code,
            };

            var jToken = this.Data as JToken;
            if (jToken != null)
            {
                result.Data = jToken.ToObject<TData>();
            }
            else if (this.Data != null)
            {
                result.Data = Converter.Cast<TData>(this.Data);
            }
            return result;
        }

        /// <summary>
        /// 从Json解析得到
        /// </summary>
        /// <param name="json">json数据源</param>
        /// <returns></returns>
        public static ApiResult Parse(string json)
        {
            return Newtonsoft.Json.JsonConvert.DeserializeObject<ApiResult>(json);
        }

        /// <summary>
        /// 从Json解析得到
        /// </summary>
        /// <typeparam name="TData">Data属性的类型</typeparam>
        /// <param name="json">json</param>
        /// <returns></returns>
        public static ApiResult<TData> Parse<TData>(string json)
        {
            return Newtonsoft.Json.JsonConvert.DeserializeObject<ApiResult<TData>>(json);
        }
    }


    /// <summary>
    /// 表示Rest风格Api结果
    /// </summary>
    [DebuggerDisplay("State = {State}")]
    public class ApiResult<T>
    {
        /// <summary>
        /// 状态
        /// </summary>
        public bool State { get; set; }

        /// <summary>
        /// 状态码
        /// </summary>
        public ErrorCode Code { get; set; }

        /// <summary>
        /// 提示信息
        /// </summary>
        public string Msg { get; set; }

        /// <summary>
        /// 业务数据
        /// </summary>
        public T Data { get; set; }
    }
}
