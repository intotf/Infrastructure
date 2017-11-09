using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace Infrastructure.Res
{
    /// <summary>
    /// 表示文件的PostedFile包装类
    /// </summary>
    public class PostedFile : HttpPostedFileBase
    {
        /// <summary>
        /// 文件名
        /// </summary>
        private readonly string fileName;

        /// <summary>
        /// 文件流
        /// </summary>
        private readonly FileStream fileStream;

        /// <summary>
        /// 文件的PostedFile包装
        /// </summary>
        /// <param name="fileName">文件名</param>
        public PostedFile(string fileName)
        {
            this.fileName = Path.GetFileName(fileName);
            this.fileStream = new FileStream(fileName, FileMode.Open);
        }

        public override int ContentLength
        {
            get
            {
                return (int)this.fileStream.Length;
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
                return this.fileName;
            }
        }

        public override Stream InputStream
        {
            get
            {
                return this.fileStream;
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="filename"></param>
        public override void SaveAs(string filename)
        {
            throw new NotSupportedException();
        }
    }
}
