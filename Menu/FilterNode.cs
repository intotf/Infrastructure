using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infrastructure.Menu
{
    /// <summary>
    /// 表示过滤保留的菜单节点
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class FilterNode<T> where T : struct
    {
        /// <summary>
        /// 获取或设置节点配置的操作行为枚举
        /// </summary>
        public T ActionEnum { get; set; }

        /// <summary>
        /// 获取或设置节点的MD5值
        /// </summary>
        public string HashMd5 { get; set; }
    }
}
