using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infrastructure.Menu
{
    /// <summary>
    /// 表示许可的菜单节点
    /// </summary>
    public sealed class PermissionNode<T> where T : struct
    {
        /// <summary>
        /// 获取或设置节点的MD5值
        /// </summary>
        public string HashMd5 { get; set; }

        /// <summary>
        /// 获取或设置节点允许的操作行为枚举
        /// </summary>
        public T AllowedAction { get; set; }
    }
}
