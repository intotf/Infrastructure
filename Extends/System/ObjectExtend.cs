using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace System
{
    /// <summary>
    /// Object扩展
    /// </summary>
    public static partial class ObjectExtend
    {
        /// <summary>
        /// 强制转换为目标类型
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <returns></returns>
        public static T OfType<T>(this object source)
        {
            return (T)source;
        }
    }
}
