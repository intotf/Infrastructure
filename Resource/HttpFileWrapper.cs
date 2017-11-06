using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace Infrastructure.Resource
{
    /// <summary>
    /// http文件包装类
    /// </summary>
    public class HttpFileWrapper : HttpPostedFileBase
    {
        private readonly string _fileName;

        private readonly MemoryStream _inputStream;

        /// <summary>
        /// http文件包装类
        /// </summary>
        /// <param name="inputStream">文件流</param>
        /// <param name="fileName">文件名</param>
        public HttpFileWrapper(MemoryStream inputStream, string fileName)
        {
            this._inputStream = inputStream;
            this._fileName = fileName;
        }

        public override int ContentLength
        {
            get
            {
                return (int)this._inputStream.Length;
            }
        }

        public override string ContentType
        {
            get
            {
                return string.Empty;
            }
        }

        public override string FileName
        {
            get
            {
                return this._fileName;
            }
        }

        public override Stream InputStream
        {
            get
            {
                return this._inputStream;
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="filename"></param>
        public override void SaveAs(string filename)
        {
            var buffer = this._inputStream.ToArray();
            File.WriteAllBytes(filename, buffer);
        }
    }
}
