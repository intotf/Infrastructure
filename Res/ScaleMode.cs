using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Infrastructure.Res
{
    /// <summary>
    /// 缩放模式
    /// </summary>
    public enum ScaleMode
    {
        /// <summary>
        /// 使X轴铺满，Y轴保持不变
        /// </summary>
        Fill_X,
        /// <summary>
        /// 使Y轴铺满，X轴保持不变
        /// </summary>
        Fill_Y,
        /// <summary>
        /// 使XY都轴铺满
        /// </summary>
        Fill_XY,
        /// <summary>
        /// 使X轴铺满，Y轴跟着比例缩放
        /// </summary>
        Fill_X_WithY,
        /// <summary>
        /// 使Y轴铺满，X轴跟着比例缩放
        /// </summary>
        Fill_Y_WithX,
        /// <summary>
        /// 使X或Y轴尽量铺满，两轴缩放比例一样
        /// </summary>
        Fill_XY_Auto,

        /// <summary>
        /// 使X和Y以最小边等比铺满，两轴缩放比较一样
        /// </summary>
        Fill_XY_Center,
    }
}
