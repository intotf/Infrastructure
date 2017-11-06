using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Infrastructure.Configs
{
    /// <summary>
    /// Http服务配置项
    /// </summary>
    public class HttpConfig : ConfigurationSection
    {
        /// <summary>
        /// 主机
        /// </summary>
        [ConfigurationProperty("Host", DefaultValue = "localhost")]
        public string Host
        {
            get
            {
                return (string)base["Host"];
            }
            set
            {
                base["Host"] = value;
            }
        }

        /// <summary>
        /// 授权码
        /// </summary>
        [ConfigurationProperty("AuthCode", DefaultValue = "tc")]
        public string AuthCode
        {
            get
            {
                return base["AuthCode"].ToString();
            }
            set
            {
                base["AuthCode"] = value;
            }
        }
    }
}
