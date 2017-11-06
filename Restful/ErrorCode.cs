using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace System.Web
{
    /// <summary>
    /// 错误码
    /// </summary>
    public enum ErrorCode
    {
        /// <summary>
        /// 无错误
        /// 0
        /// </summary>
        NoError = 0,

        /// <summary>
        /// 请求的Token无效
        /// 1
        /// </summary>
        [Display(Name = "请求的Token无效")]
        TokenInvalid = 1,

        /// <summary>
        /// 账号已被禁用
        /// 2
        /// </summary>
        [Display(Name = "账号已被禁用")]
        Disabled = 2,

        /// <summary>
        /// 设备不存在
        /// 3
        /// </summary>
        [Display(Name = "设备不存在")]
        EQNonentity = 3,

        /// <summary>
        /// 账号或密码不正确
        /// 4
        /// </summary>
        [Display(Name = "账号或密码不正确")]
        AccountOrPasswordIsNotCorrect = 4,

        /// <summary>
        /// 未绑定的第三方账号
        /// 5
        /// </summary>
        [Display(Name = "未绑定的第三方账号")]
        ThirdAccountIdNoBinding = 5,
    }
}