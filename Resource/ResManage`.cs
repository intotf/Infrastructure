using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace Infrastructure.Resource
{
    /// <summary>
    /// 资源文件统一管理类
    /// 提供对资源文件的保存、分析和其它常用操作
    /// 支持可配置化是否以月份分隔资源文件   
    /// </summary>
    public sealed class ResManage<T> where T : struct
    {
        /// <summary>
        /// 获取资源的类型
        /// </summary>
        public T ResType { get; private set; }

        /// <summary>
        /// 获取资源的短文件名
        /// 对应保存到数据库的文件名
        /// </summary>
        public string FileName { get; private set; }

        /// <summary>
        /// 获取资源的完整文件名
        /// </summary>
        public string FullFileName
        {
            get
            {
                if (this.FileName.IsNullOrEmpty() == true)
                {
                    return string.Empty;
                }
                return Path.Combine(ResManage.RootFullPath, this.ResType.ToString(), this.FileName);
            }
        }
        /// <summary>
        /// 获取资源的web相对路径
        /// </summary>
        public string UrlName
        {
            get
            {
                if (this.FileName.IsNullOrEmpty() == true)
                {
                    return string.Empty;
                }
                return string.Concat(ResManage.RootURL, this.ResType.ToString(), "/", this.FileName.Replace(@"\", "/"));
            }
        }

        /// <summary>
        /// 获取资源的完整web路径
        /// </summary>
        public string FullUrlName
        {
            get
            {
                if (this.FileName.IsNullOrEmpty() == true)
                {
                    return string.Empty;
                }
                return string.Concat(ResManage.RootFullURL, this.ResType.ToString(), "/", this.FileName.Replace(@"\", "/"));
            }
        }

        /// <summary>
        /// 获取资源文件是否存在
        /// </summary>
        public bool FileExist
        {
            get
            {
                return File.Exists(this.FullFileName);
            }
        }

        /// <summary>
        /// 构造器
        /// </summary>
        /// <param name="fileName">文件本地或web路径(包含ResType)</param>
        internal ResManage(string fileName)
        {
            if (fileName != null)
            {
                fileName = HttpUtility.UrlDecode(fileName);
                var inputString = fileName.Replace("/", @"\");

                this.FileName = this.GetFileName(inputString);
                this.ResType = this.GetResType(inputString);
            }
        }

        /// <summary>
        /// 构造器
        /// </summary>
        /// <param name="fileName">文件本地或web路径(包含或不包含ResType)</param>
        /// <param name="resType">资源类型</param>      
        internal ResManage(string fileName, T resType)
        {
            this.ResType = resType;

            fileName = HttpUtility.UrlDecode(fileName);
            var inputString = fileName.Replace("/", @"\");
            this.FileName = this.GetFileName(inputString);
        }

        /// <summary>
        /// 分析和获取文件名
        /// </summary>
        /// <param name="input">输入的文件路径信息</param>    
        private string GetFileName(string input)
        {
            var month = @"\d{4}_\d{2}";
            var guid = @"\w{8}_\w{4}_\w{4}_\w{4}_\w{12}";
            var name = @".+?";
            var ext = @"\.\w{2,4}$";

            var guidFloder = string.Format(@"{0}\\{1}{2}", guid, name, ext);
            var guidNoFloder = string.Format(@"{0}{1}", guid, ext);

            var monthGuidFloder = string.Format(@"{0}\\{1}", month, guidFloder);
            var monthGuidNoFloder = string.Format(@"{0}\\{1}", month, guidNoFloder);

            var fileNameReg = string.Join("|", monthGuidFloder, monthGuidNoFloder, guidFloder, guidNoFloder);
            return input.Match(fileNameReg);
        }

        /// <summary>
        /// 分析和获取资源类型
        /// </summary>
        /// <param name="input">输入的文件路径信息</param>
        private T GetResType(string input)
        {
            var resType = default(T);
            if (string.IsNullOrEmpty(this.FileName) == false)
            {
                var value = input.Replace(this.FileName, string.Empty).Match(@"\w+\\$").Replace(@"\", string.Empty).ToLower();
                Enum.TryParse<T>(value, true, out resType);
            }
            return resType;
        }

        /// <summary>
        /// 删除包含的资源文件
        /// 不存在不会报错
        /// </summary>
        public bool Delete()
        {
            try
            {
                File.Delete(this.FullFileName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 复制到指定类型资源的路径中
        /// 如果指定文件存在，则覆盖
        /// 成功则返回true
        /// </summary>
        /// <param name="targetResType">目标类型</param>
        /// <returns></returns>
        public ResManage<T> CopyTo(T targetResType)
        {
            if (this.FileExist == false)
            {
                throw new FileNotFoundException("文件不存在", this.FullFileName);
            }

            if (targetResType.Equals(this.ResType) == true)
            {
                throw new ArgumentException("resType参数有误", "resType");
            }

            var match = Regex.Match(this.FileName, @"$\d{4}_\d{2}\\");
            var fileName = Regex.Replace(this.FileName, @"$\d{4}_\d{2}\\", string.Empty);

            var resTypeEnum = (Enum)Enum.Parse(typeof(T), targetResType.ToString());
            if (resTypeEnum.IsDefined<ResSplitByMonthAttribute>() == true)
            {
                fileName = string.Format(@"{0}\{1}", match.Success ? match.Value : DateTime.Now.ToString("yyyy_MM"), fileName);
            }

            var res = ResManage.Parse(fileName, targetResType);
            Directory.CreateDirectory(Path.GetDirectoryName(res.FullFileName));
            File.Copy(this.FullFileName, res.FullFileName, true);
            return res;
        }

        /// <summary>
        /// 移动到指定类型资源的路径中
        /// 如果指定文件存在，则覆盖
        /// 返回新目标资源
        /// </summary>
        /// <param name="targetResType">目标类型</param>
        /// <returns></returns>
        public ResManage<T> MoveTo(T targetResType)
        {
            if (this.ResType.Equals(targetResType) == true)
            {
                return this;
            }
            var targetRes = this.CopyTo(targetResType);
            this.Delete();
            return targetRes;
        }

        /// <summary>
        /// 移动到指定类型资源的路径中
        /// 如果指定文件存在，则覆盖
        /// 返回新目标资源
        /// </summary>
        /// <param name="resType">当资源类型为resType时才移动</param>
        /// <param name="targetResType">目标类型</param>
        /// <returns></returns>
        public ResManage<T> MoveToWhen(T resType, T targetResType)
        {
            if (this.ResType.Equals(resType) == false)
            {
                return this;
            }
            return this.MoveTo(targetResType);
        }

        /// <summary>
        /// 重命名资源文件名
        /// 成功则返回true
        /// </summary>
        /// <param name="newName">新文件名</param>
        /// <returns></returns>
        public bool ReName(string newName)
        {
            if (this.FileExist == false)
            {
                return false;
            }

            newName = Path.GetFileNameWithoutExtension(newName.NullThenEmpty());
            if (string.IsNullOrEmpty(newName) == true)
            {
                return false;
            }

            var dir = Path.GetDirectoryName(this.FileName);
            newName = string.Concat(newName, Path.GetExtension(this.FileName));

            var sourceFile = this.FullFileName;
            this.FileName = Path.Combine(dir, newName);
            File.Move(sourceFile, this.FullFileName);
            return true;
        }

        /// <summary>
        /// 转换为保存到数据库的文件名
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return this.FileName;
        }
    }
}
