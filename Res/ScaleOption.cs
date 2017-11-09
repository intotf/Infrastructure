using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Infrastructure.Res
{
    /// <summary>
    /// 缩放选项
    /// </summary>
    public enum ScaleOption
    {
        /// <summary>
        /// 总是缩放
        /// </summary>
        AllWay,
        /// <summary>
        /// 当图片尺寸比目标小时不缩放
        /// </summary>
        NoSacleWhenSmall
    }
}
