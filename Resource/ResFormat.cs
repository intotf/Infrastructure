using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Infrastructure.Resource
{
    /// <summary>
    /// 资源格式
    /// </summary>
    public sealed class ResFormat
    {
        /// <summary>
        /// 格式
        /// </summary>
        public string Extension { get; private set; }

        /// <summary>
        /// 格式标识
        /// </summary>
        public byte[] FormatValue { get; private set; }

        /// <summary>
        /// 文件格式
        /// </summary>
        /// <param name="ext">格式</param>
        /// <param name="formatValue">格式标识</param>
        private ResFormat(string ext, byte[] formatValue)
        {
            this.Extension = ext.ToLower();
            this.FormatValue = formatValue;
        }

        /// <summary>
        /// 获取哈希码
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return this.Extension.NullThenEmpty().GetHashCode() ^ this.FormatValue[0].GetHashCode() ^ this.FormatValue[1].GetHashCode();
        }

        /// <summary>
        /// 比较是否相等
        /// </summary>
        /// <param name="obj">目标对象</param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }

            return this.GetHashCode() == obj.GetHashCode();
        }

        /// <summary>
        /// 流是否为指定的格式一种
        /// </summary>
        /// <param name="stream">流</param>
        /// <param name="exts">格式</param>
        /// <returns></returns>
        public static bool IsInFormat(Stream stream, string exts)
        {
            if (stream == null)
            {
                return false;
            }

            var extArray = exts.Matches(@"\.\w+").Select(item => item.ToLower()).ToArray();
            if (extArray.Length == 0)
            {
                return false;
            }

            var formats = from k in ResFormat.KNOW_FORMAT_LIST
                          join e in extArray
                          on k.Extension equals e
                          select k;

            var formatValue = ResFormat.ReadStreamFormatValue(stream);

            return formats.Any(f => formatValue.SequenceEqual(f.FormatValue));
        }


        /// <summary>
        /// 匹配文件格式是否属于扩展名
        /// </summary>
        /// <param name="ext">文件名或文件扩展名</param>
        /// <param name="stream">文件流</param>
        /// <param name="allowExts">允许的扩展名(.txt;.doc)</param>
        /// <returns></returns>
        public static bool IsFormat(string ext, Stream stream, string allowExts)
        {
            if (stream == null || ext.IsNullOrEmpty() == true)
            {
                return false;
            }


            var extArray = allowExts.Matches(@"\.\w+").Select(item => item.ToLower()).ToArray();
            if (extArray.Length == 0)
            {
                return false;
            }

            var formatValue = ResFormat.ReadStreamFormatValue(stream);
            if (formatValue == null)
            {
                return false;
            }

            var format = new ResFormat(Path.GetExtension(ext).ToLower(), formatValue);

            var formats = from k in ResFormat.KNOW_FORMAT_LIST
                          join e in extArray
                          on k.Extension equals e
                          select k;

            return formats.Any(f => f.Equals(format));
        }

        /// <summary>
        /// 获取流的前二个字节
        /// </summary>
        /// <param name="stream">流</param>
        /// <returns></returns>
        private static byte[] ReadStreamFormatValue(Stream stream)
        {
            if (stream == null || stream.Length < 2)
            {
                return null;
            }

            stream.Position = 0;
            var buffer = new byte[2];
            stream.Read(buffer, 0, buffer.Length);
            return buffer;
        }


        /// <summary>
        /// 常用已知的文件格式信息
        /// </summary>
        private static List<ResFormat> KNOW_FORMAT_LIST = new List<ResFormat>()
        {
            new ResFormat(".zip", new byte[] { 0x50, 0x4B }),
            new ResFormat(".rar", new byte[] { 0x52, 0x61 }),

            new ResFormat(".bmp", new byte[] { 0x42, 0x4D }),
            new ResFormat(".png", new byte[] { 0x89, 0x50 }),
            new ResFormat(".jpg", new byte[] { 0xFF, 0xD8 }),
            new ResFormat(".jpeg", new byte[] { 0xFF, 0xD8 }),
            new ResFormat(".gif", new byte[] { 0x47, 0x49 }),

            new ResFormat(".doc", new byte[] { 0xD0, 0xCF }),
            new ResFormat(".ppt", new byte[] { 0xD0, 0xCF }),
            new ResFormat(".xls", new byte[] { 0xD0, 0xCF }),
            new ResFormat(".docx", new byte[] { 0x50, 0x4B }),
            new ResFormat(".pptx", new byte[] { 0x50, 0x4B }),
            new ResFormat(".xlsx", new byte[] { 0x50, 0x4B }),
            new ResFormat(".pdf", new byte[] { 0x25, 0x50 }),
            new ResFormat(".rtf", new byte[] { 0x7B, 0x5C }),
          
            new ResFormat(".swf", new byte[] { 0x43, 0x57 }),
            new ResFormat(".flv", new byte[] { 0x46, 0x4C })
        };
    }
}
