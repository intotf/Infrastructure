using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Infrastructure.Configs
{
    /// <summary>
    /// 表示Redis配置
    /// </summary>
    public class RedisConfig : ConfigurationSection
    {
        /// <summary>
        /// 域名
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
        /// 写
        /// </summary>
        [ConfigurationProperty("MaxWritePoolSize", DefaultValue = 150)]
        public int MaxWritePoolSize
        {
            get
            {
                return (int)base["MaxWritePoolSize"];
            }
            set
            {
                base["MaxWritePoolSize"] = value;
            }
        }

        /// <summary>
        /// 读
        /// </summary>
        [ConfigurationProperty("MaxReadPoolSize", DefaultValue = 150)]
        public int MaxReadPoolSize
        {
            get
            {
                return (int)base["MaxReadPoolSize"];
            }
            set
            {
                base["MaxReadPoolSize"] = value;
            }
        }

        /// <summary>
        /// 过期时间
        /// </summary>
        [ConfigurationProperty("ExpireMinutes", DefaultValue = 30)]
        public int ExpireMinutes
        {
            get
            {
                return (int)base["ExpireMinutes"];
            }
            set
            {
                base["ExpireMinutes"] = value;
            }
        }

        /// <summary>
        /// 数据库
        /// </summary>
        [ConfigurationProperty("Database", DefaultValue = 0)]
        public int Database
        {
            get
            {
                return (int)base["Database"];
            }
            set
            {
                base["Database"] = value;
            }
        }
    }
}
