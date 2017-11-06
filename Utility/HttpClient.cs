using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Web;
using System.IO;
using System.Collections.Concurrent;
using Infrastructure.Reflection;

namespace Infrastructure.Utility
{
    /// <summary>
    /// HttpClient
    /// </summary>
    public class HttpClient : WebClient
    {
        /// <summary>
        /// 获取模型的参数
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="model">模型</param>
        /// <returns></returns>
        private static string GetParameter<T>(T model) where T : class
        {
            if (model == null)
            {
                return string.Empty;
            }

            var items = model.GetType().GetProperties().Select(item =>
            {
                var value = item.GetValue(model, null);
                var valueString = value == null ? string.Empty : value.ToString();
                return string.Format("{0}={1}", item.Name, HttpUtility.UrlEncode(valueString));
            });
            return string.Join("&", items.ToArray());
        }

        /// <summary>
        /// 获取模型的参数
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="model">模型</param>
        /// <returns></returns>
        public static string GetParameter<T>(params T[] model) where T : class
        {
            if (model == null)
            {
                return string.Empty;
            }
            var items = model.Select(item => HttpClient.GetParameter(item));
            return string.Join("&", items);
        }


        /// <summary>
        /// 获取或设置超时间（毫秒）
        /// 默认60秒
        /// </summary>
        public int TimeOut { get; set; }


        /// <summary>
        /// 获取或设置Cookie容器
        /// </summary>
        public CookieContainer CookieContainer { get; set; }

        /// <summary>
        /// http客户端
        /// </summary>
        public HttpClient()
        {
            this.Proxy = null;
            this.TimeOut = 60 * 1000;
        }


        /// <summary>
        /// 为指定资源返回一个 System.Net.WebRequest 对象
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        protected override WebRequest GetWebRequest(Uri address)
        {
            var request = base.GetWebRequest(address) as HttpWebRequest;
            request.CookieContainer = this.CookieContainer;
            if (this.TimeOut > 0)
            {
                request.Timeout = this.TimeOut;
            }
            return request;
        }

        /// <summary>
        /// Get请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <returns></returns>
        public byte[] HttpGet(string address)
        {
            return this.DownloadData(address);
        }

        /// <summary>
        /// Get请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="encoding">编码</param>
        /// <returns></returns>
        public string HttpGet(string address, Encoding encoding)
        {
            var bytes = this.HttpGet(address);
            return encoding.GetString(bytes);
        }


        /// <summary>
        /// Get请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="parameters">参数</param>
        /// <param name="encoding">编码</param>
        /// <returns></returns>
        public string HttpGet<T>(string address, T parameter, Encoding encoding) where T : class
        {
            var param = HttpClient.GetParameter(parameter);
            if (param == null)
            {
                return this.HttpGet(address, encoding);
            }

            if (address.Contains('?'))
            {
                address = address + "&" + param;
            }
            else
            {
                address = address + "?" + param;
            }
            return this.HttpGet(address, encoding);
        }


        /// <summary>
        /// Get请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="parameters">参数</param>
        /// <param name="encoding">编码</param>       
        /// <returns></returns>
        public string HttpGet<T>(string address, T[] parameters, Encoding encoding) where T : class
        {
            var param = HttpClient.GetParameter(parameters);
            if (param == null)
            {
                return this.HttpGet(address, encoding);
            }

            if (address.Contains('?'))
            {
                address = address + "&" + param;
            }
            else
            {
                address = address + "?" + param;
            }
            return this.HttpGet(address, encoding);
        }

        /// <summary>
        /// Post请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="parameter">参数</param>
        /// <returns></returns>
        public byte[] HttpPost(string address, byte[] parameter)
        {
            if (string.IsNullOrEmpty(this.Headers[HttpRequestHeader.ContentType]))
            {
                this.Headers.Add(HttpRequestHeader.ContentType, "application/x-www-form-urlencoded");
            }
            return this.UploadData(address, "POST", parameter);
        }

        /// <summary>
        /// Post请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="parameter">参数</param>
        /// <param name="encoding">编码</param>
        /// <returns></returns>
        public string HttpPost(string address, byte[] parameter, Encoding encoding)
        {
            var bytes = this.HttpPost(address, parameter);
            return encoding.GetString(bytes);
        }

        /// <summary>
        /// Post请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="parameter">用&连接的参数</param>
        /// <param name="encoding">编码</param>
        /// <returns></returns>
        public string HttpPost(string address, string parameter, Encoding encoding)
        {
            var bytes = encoding.GetBytes(parameter);
            return this.HttpPost(address, bytes, encoding);
        }

        /// <summary>
        /// Post请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="parameters">参数</param>
        /// <param name="encoding">编码</param>
        /// <returns></returns>
        public string HttpPost<T>(string address, T parameters, Encoding encoding) where T : class
        {
            var param = HttpClient.GetParameter(parameters);
            return this.HttpPost(address, param, encoding);
        }

        /// <summary>
        /// Post请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="parameters">参数</param>
        /// <param name="encoding">编码</param>
        /// <returns></returns>
        public string HttpPost<T>(string address, T[] parameters, Encoding encoding) where T : class
        {
            var param = HttpClient.GetParameter(parameters);
            return this.HttpPost(address, param, encoding);
        }

        /// <summary>
        /// Post请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="form">表单</param>
        /// <returns></returns>
        public byte[] HttpPost(string address, MultipartForm form)
        {
            this.Headers.Add(HttpRequestHeader.ContentType, "multipart/form-data; boundary=----" + form.Boundary);
            return this.UploadData(address, "POST", form.ToArray());
        }

        /// <summary>
        /// Post请求
        /// </summary>
        /// <param name="address">地址</param>
        /// <param name="form">表单</param>
        /// <param name="encoding">解码</param>
        /// <returns></returns>
        public string HttpPost(string address, MultipartForm form, Encoding encoding)
        {
            var bytes = this.HttpPost(address, form);
            return encoding.GetString(bytes);
        }

        /// <summary>
        /// 表示二进制表单
        /// </summary>
        public class MultipartForm : IDisposable
        {
            private readonly byte[] firstBoundary;

            private readonly byte[] lastBoundary;

            private readonly MemoryStream stream = new MemoryStream();

            /// <summary>
            /// 获取分割线
            /// </summary>
            public string Boundary { get; private set; }

            /// <summary>
            /// 二进制表单
            /// </summary>
            public MultipartForm()
            {
                this.Boundary = "Boundary" + Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
                this.firstBoundary = Encoding.UTF8.GetBytes(string.Format("------{0}\r\n", this.Boundary));
                this.lastBoundary = Encoding.UTF8.GetBytes(string.Format("------{0}--\r\n", this.Boundary));
            }

            /// <summary>
            /// 添加实体
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="model">实例</param>
            /// <returns></returns>
            public MultipartForm Add<T>(T model) where T : class
            {
                if (model == null)
                {
                    return this;
                }
                var properties = Property.GetProperties(typeof(T));
                foreach (var item in properties)
                {
                    var value = item.GetValue(model);
                    this.Add(item.Name, value == null ? null : value.ToString());
                }
                return this;
            }

            /// <summary>
            /// 添加参数
            /// </summary>
            /// <param name="name">名称</param>
            /// <param name="value">值</param>
            /// <returns></returns>
            public MultipartForm Add(string name, IEnumerable<string> value)
            {
                if (value != null)
                {
                    foreach (var item in value)
                    {
                        this.Add(name, item);
                    }
                }
                return this;
            }

            /// <summary>
            /// 添加参数
            /// </summary>
            /// <param name="name">名称</param>
            /// <param name="value">值</param>
            /// <returns></returns>
            public MultipartForm Add(string name, string value)
            {
                var text = string.Format("Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}\r\n", name, value);
                var bytes = Encoding.UTF8.GetBytes(text);
                this.stream.Write(this.firstBoundary, 0, this.firstBoundary.Length);
                this.stream.Write(bytes, 0, bytes.Length);
                return this;
            }

            /// <summary>
            /// 添加文件
            /// </summary>
            /// <param name="name">名称</param>
            /// <param name="file">文件</param>
            /// <returns></returns>
            public MultipartForm AddFile(string name, HttpPostedFileBase file)
            {
                if (file == null)
                {
                    return this;
                }
                var bytes = new byte[file.InputStream.Length];
                file.InputStream.Read(bytes, 0, bytes.Length);
                return this.AddFile(name, bytes, file.FileName, file.ContentType);
            }

            /// <summary>
            /// 添加文件
            /// </summary>
            /// <param name="name">名称</param>
            /// <param name="filePath">文件本地完整路径</param>
            /// <returns></returns>
            public MultipartForm AddFile(string name, string filePath)
            {
                if (filePath.IsNullOrEmpty() || File.Exists(filePath) == false)
                {
                    return this;
                }

                using (var stream = new FileStream(filePath, FileMode.Open))
                {
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, bytes.Length);
                    return this.AddFile(name, bytes, filePath, null);
                }
            }

            /// <summary>
            /// 添加文件
            /// </summary>
            /// <param name="name">名称</param>
            /// <param name="buffer">文件内容</param>
            /// <param name="fileName">文件名</param>
            /// <param name="contentType">文件类型</param>
            /// <returns></returns>
            public MultipartForm AddFile(string name, byte[] buffer, string fileName, string contentType)
            {
                if (string.IsNullOrEmpty(contentType))
                {
                    contentType = "application/octet-stream";
                }
                var text = string.Format("Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: {2}\r\n\r\n", name, fileName, contentType);
                var bytes = Encoding.UTF8.GetBytes(text);
                this.stream.Write(this.firstBoundary, 0, this.firstBoundary.Length);
                this.stream.Write(bytes, 0, bytes.Length);
                this.stream.Write(buffer, 0, buffer.Length);
                var crlf = Encoding.UTF8.GetBytes("\r\n");
                this.stream.Write(crlf, 0, crlf.Length);
                return this;
            }

            /// <summary>
            /// 转换为数组
            /// </summary>
            /// <returns></returns>
            public byte[] ToArray()
            {
                this.stream.Write(this.lastBoundary, 0, this.lastBoundary.Length);
                this.stream.Position = this.stream.Position - this.lastBoundary.Length;
                return this.stream.ToArray();
            }

            /// <summary>
            /// 释放资源
            /// </summary>
            public void Dispose()
            {
                this.stream.Dispose();
            }
        }
    }
}
