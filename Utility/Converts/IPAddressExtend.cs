using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using Infrastructure.Utility.Converts;

namespace Infrastructure.Utility.Converts
{
    /// <summary>
    /// IP地址扩展
    /// </summary>
    public static class IPAddressExtend
    {
        /// <summary>
        /// 将IP地址转换为数字
        /// </summary>
        /// <param name="ip"></param>
        /// <returns></returns>
        public static int ToInt32Big(this IPAddress ip)
        {
            var bytes = ip.GetAddressBytes();
            return ByteConverter.ToInt32(bytes, 0, Endians.Big);
        }

        /// <summary>
        /// 在当前IP地址+N 返回新的IP地址
        /// </summary>
        /// <param name="ip">当前IP地址</param>
        /// <param name="value">增加数</param>
        /// <returns></returns>
        public static IPAddress Add(this IPAddress ip, int value)
        {
            var bytes = ByteConverter.ToBytes(ip.ToInt32Big() + value, Endians.Big);
            return new IPAddress(bytes);
        }

        /// <summary>
        /// 判断IP 是否可用，排除 255和1
        /// </summary>
        /// <param name="ip"></param>
        /// <returns></returns>
        public static bool IsValid(this IPAddress ip)
        {
            var last = ip.GetAddressBytes().LastOrDefault();
            return last != byte.MaxValue && last != byte.MinValue;
        }

        /// <summary>
        /// 获取下一下Ip地址
        /// </summary>
        /// <param name="ip"></param>
        /// <returns></returns>
        public static IPAddress GetNextValid(this IPAddress ip)
        {
            while (true)
            {
                ip = ip.Add(1);
                if (ip.IsValid() == true)
                {
                    return ip;
                }
            }
        }

        /// <summary>
        /// 字符串转IPAddress
        /// </summary>
        /// <param name="ip"></param>
        /// <returns></returns>
        public static IPAddress ToIPAddress(this string ip)
        {
            if (!string.IsNullOrEmpty(ip))
            {
                return IPAddress.Parse(ip);
            }
            return null;
        }
    }
}
