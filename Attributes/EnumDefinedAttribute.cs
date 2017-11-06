using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infrastructure.Attributes
{
    /// <summary>
    /// 表示验证模型枚举值是否定义特性
    /// </summary>
    public class EnumDefinedAttribute : System.ComponentModel.DataAnnotations.ValidationAttribute
    {
        public override bool IsValid(object value)
        {
            if (value == null)
            {
                return true;
            }
            return Enum.IsDefined(value.GetType(), value);
        }
    }
}
