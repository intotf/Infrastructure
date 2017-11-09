using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infrastructure.Attributes
{
    /// <summary>
    /// 表示内部使用的字段或属性
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Method | AttributeTargets.Property | AttributeTargets.Class | AttributeTargets.Interface | AttributeTargets.Enum)]
    public class InternalAttribute : Attribute
    {
    }
}
