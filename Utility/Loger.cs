using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace Infrastructure.Utility
{
    /// <summary>
    /// 日志
    /// </summary>
    public class Loger
    {
        /// <summary>
        /// 获取今天的日志文件
        /// </summary>
        public static Loger Today(string name = "Log")
        {
            var fileName = string.Format("{0}_{1}", name, DateTime.Today.ToString("yyyyMMdd"));
            return new Loger(fileName);
        }

        /// <summary>
        /// 获取现在的日志文件
        /// </summary>
        public static Loger Now(string name = "Log")
        {
            var fileName = string.Format("{0}_{1}", name, DateTime.Now.ToString("yyyyMMddHHmmssfff"));
            return new Loger(fileName);
        }

        /// <summary>
        /// 日志文件同步锁
        /// </summary>
        private static readonly object syncRoot = new object();

        /// <summary>
        /// 文件路径
        /// </summary>
        private readonly string fileFullName;

        /// <summary>
        /// 内容
        /// </summary>
        private readonly StringBuilder sb = new StringBuilder();

        /// <summary>
        /// 日志
        /// </summary>
        /// <param name="fileName">文件名</param>
        public Loger(string fileName)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException();
            }

            if (Path.GetExtension(fileName).IsNullOrEmpty())
            {
                fileName = fileName + ".txt";
            }

            if (HttpContext.Current != null)
            {
                fileName = HttpContext.Current.Server.MapPath("~/" + fileName);
            }

            var dir = Path.GetDirectoryName(fileName);
            if (dir.IsNullOrEmpty() == false)
            {
                Directory.CreateDirectory(dir);
            }
            this.fileFullName = fileName;
        }

        /// <summary>
        /// 追加日志
        /// </summary>
        /// <param name="log">日志</param>
        /// <returns></returns>
        public Loger Append(object log)
        {
            if (log != null)
            {
                this.sb.Append(log.ToString());
            }
            return this;
        }

        /// <summary>
        /// 追加日志
        /// </summary>
        /// <param name="log">日志</param>
        /// <param name="args">参数</param>
        /// <returns></returns>
        public Loger Append(object log, params object[] args)
        {
            if (log != null)
            {
                var logTxt = string.Format(log.ToString(), args);
                this.sb.Append(logTxt);
            }
            return this;
        }

        /// <summary>
        /// 追加日志
        /// </summary>
        /// <param name="log">日志</param>
        /// <param name="args">参数</param>
        /// <returns></returns>
        public Loger AppendLine()
        {
            this.sb.AppendLine();
            return this;
        }

        /// <summary>
        /// 追加日志
        /// </summary>
        /// <param name="log">日志</param>
        /// <returns></returns>
        public Loger AppendLine(object log)
        {
            if (log != null)
            {
                this.sb.AppendLine(log.ToString());
            }
            return this;
        }

        /// <summary>
        /// 追加日志
        /// </summary>
        /// <param name="log">日志</param>
        /// <param name="args">参数</param>
        /// <returns></returns>
        public Loger AppendLine(object log, params object[] args)
        {
            if (log != null)
            {
                var logTxt = string.Format(log.ToString(), args);
                this.sb.AppendLine(logTxt);
            }
            return this;
        }

        /// <summary>
        /// 追加日志
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="model">模型</param>
        /// <returns></returns>
        public Loger AppendLineAsJson<T>(T model)
        {
            var json = JsonSerializer.Serialize(model);
            return this.AppendLine(json);
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            lock (syncRoot)
            {
                File.AppendAllText(this.fileFullName, this.sb.ToString());
                sb.Clear();
            }
        }
    }
}
