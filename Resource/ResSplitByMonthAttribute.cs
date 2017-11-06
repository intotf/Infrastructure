using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infrastructure.Resource
{
    /// <summary>
    /// 表示此类资源将按月分开存储的特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Field, AllowMultiple = false)]
    public sealed class ResSplitByMonthAttribute : Attribute
    {
    }
}
