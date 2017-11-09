using System;
using System.Collections.Generic;
using System.Runtime.Caching;
using System.Text;
using System.Threading.Tasks;

namespace HDP.Common.Utility
{
    /// <summary>
    /// 表示绝对过期的内存缓存
    /// </summary>
    /// <typeparam name="T">缓存类型</typeparam>
    public class AbsoluteCache<T>
    {
        /// <summary>
        /// 同步锁
        /// </summary>
        public readonly object SyncRoot = new object();

        /// <summary>
        /// 获取或设置绝对过期时间
        /// </summary>
        public TimeSpan Expiration { get; private set; }


        /// <summary>
        /// 绝对过期的内存缓存
        /// </summary>
        /// <param name="absoluteExpiration">从添加开始，到过期的时间戳</param>
        public AbsoluteCache(TimeSpan absoluteExpiration)
        {
            this.Expiration = absoluteExpiration;
        }

        /// <summary>
        /// 获取或添加缓存
        /// </summary>
        /// <param name="key">键</param>
        /// <param name="valueFactory">值</param>
        /// <returns></returns>
        public T GetOrAdd(string key, Func<string, T> valueFactory)
        {
            lock (this.SyncRoot)
            {
                if (MemoryCache.Default.Contains(key) == true)
                {
                    var itemValue = MemoryCache.Default.Get(key);
                    return itemValue == DBNull.Value ? default(T) : (T)itemValue;
                }
                else
                {
                    var value = valueFactory(key);
                    var itemValue = value == null ? DBNull.Value : (object)value;

                    MemoryCache.Default.Set(key, itemValue, new CacheItemPolicy { AbsoluteExpiration = DateTime.Now.Add(this.Expiration) });
                    return value;
                }
            }
        }

        /// <summary>
        /// 移除缓存
        /// </summary>
        /// <param name="key">键</param>
        public void Remove(string key)
        {
            lock (this.SyncRoot)
            {
                MemoryCache.Default.Remove(key);
            }
        }
    }
}
