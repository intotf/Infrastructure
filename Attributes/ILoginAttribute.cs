using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infrastructure.Attributes
{
    /// <summary>
    /// 定义登录验证特性的接口
    /// </summary>
    public interface ILoginAttribute
    {
        /// <summary>
        /// 获取需要登录的角色名
        /// </summary>
        string Role { get; }
    }
}
