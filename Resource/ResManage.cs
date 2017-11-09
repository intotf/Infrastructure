using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Reflection;
using System.Web.Mvc;
using System.Text.RegularExpressions;

namespace Infrastructure.Resource
{
    /// <summary>
    /// 资源文件统一管理类
    /// 提供对资源文件的保存、分析和其它常用操作
    /// 支持可配置化是否以月份分隔资源文件   
    /// </summary>
    public static class ResManage
    {
        /// <summary>
        /// 上传目录的文件夹名
        /// </summary>
        private readonly static string ROOT = "Res";

        /// <summary>
        /// 获取上传目录的本地完整路径
        /// </summary>
        public static string RootFullPath
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ResManage.ROOT);
            }
        }

        /// <summary>
        /// 获取上传目录的相对网络路径
        /// </summary>
        public static string RootURL
        {
            get
            {
                if (HttpContext.Current == null)
                {
                    throw new NotSupportedException();
                }

                var path = HttpContext.Current.Request.ApplicationPath;
                if (path.EndsWith("/") == false)
                {
                    path = path + "/";
                }
                return string.Concat(path, ResManage.ROOT, "/");
            }
        }
        /// <summary>
        /// 获取上传目录的完整网络路径
        /// 此属性依赖于HttpConext
        /// </summary>
        public static string RootFullURL
        {
            get
            {
                if (HttpContext.Current == null)
                {
                    throw new NotSupportedException();
                }

                var path = HttpContext.Current.Request.ApplicationPath;
                if (path.EndsWith("/") == false)
                {
                    path = path + "/";
                }
                return string.Concat(HttpContext.Current.Request.Url.Scheme, "://", HttpContext.Current.Request.Url.Authority, path, ResManage.ROOT, "/");
            }
        }

        /// <summary>
        /// 从文件名获得资源信息对象
        /// </summary>
        /// <param name="fileName">本地或web文件名[包含ResType]</param>
        /// <exception cref="ArgumentNullException">fileName</exception>
        /// <returns></returns>
        public static ResManage<T> Parse<T>(string fileName) where T : struct
        {            
            if (typeof(T).IsEnum == false)
            {
                throw new Exception("泛型参数类型必须为枚举类型");
            }
            return new ResManage<T>(fileName);
        }

        /// <summary>
        /// 从文件名获得资源信息对象
        /// </summary>
        /// <param name="fileName">本地或web文件名[包含或不包含ResType]</param>
        /// <param name="resType">指定资源类型</param>
        /// <exception cref="ArgumentNullException">fileName</exception>
        public static ResManage<T> Parse<T>(string fileName, T resType) where T : struct
        {            
            if (typeof(T).IsEnum == false)
            {
                throw new Exception("泛型参数类型必须为枚举类型");
            }
            return new ResManage<T>(fileName, resType);
        }

        /// <summary>
        /// 保存上传的文件对象到指定类型文件夹下
        /// 返回保存后资源文件的信息
        /// </summary>
        /// <param name="file">上传的文件</param>
        /// <param name="resType">资源类型</param>
        /// <param name="idName">用以作文件名或文件的父文件夹名</param>
        /// <param name="useIdFolder">是否使用idName作父文件夹包裹文件，如果为ture,文件名将不变</param>
        /// <param name="keepName">当useIdFolder为true且keepName为false时，文件名为随机名</param>
        /// <returns></returns>
        public static ResManage<T> SavePostedFile<T>(HttpPostedFileBase file, T resType, Guid idName, bool useIdFolder = false, bool keepName = false) where T : struct
        {
            if (typeof(T).IsEnum == false)
            {
                throw new Exception("泛型参数类型必须为枚举类型");
            }
            return ResManage.SaveFile(file, resType, idName, useIdFolder, keepName);
        }

        /// <summary>
        /// 保存上传的文件对象到指定类型文件夹下
        /// 返回保存后资源文件的信息
        /// </summary>
        /// <param name="file">上传的文件</param>
        /// <param name="resType">资源类型</param>
        /// <param name="idName">用以作文件名或文件的父文件夹名</param>
        /// <param name="useIdFolder">是否使用idName作父文件夹包裹文件，如果为ture,文件名将不变</param>
        /// <param name="keepName">当useIdFolder为true且keepName为false时，文件名为随机名</param>
        /// <returns></returns>
        private static ResManage<T> SaveFile<T>(HttpPostedFileBase file, T resType, object idName, bool useIdFolder = false, bool keepName = false) where T : struct
        {
            if (file == null || file.ContentLength == 0)
            {
                throw new ArgumentNullException("file");
            }

            var idNameString = idName.ToString().Replace("-", "_");
            string fileName = string.Concat(idNameString, Path.GetExtension(file.FileName));

            if (useIdFolder == true)
            {
                if (keepName == true)
                {
                    fileName = Path.Combine(idNameString, file.FileName);
                }
                else
                {
                    var randomName = string.Concat(Guid.NewGuid().ToString().Replace("-", "_"), Path.GetExtension(file.FileName));
                    fileName = Path.Combine(idNameString, randomName);
                }
            }

            var resTypeEnum = (Enum)Enum.Parse(typeof(T), resType.ToString());
            if (resTypeEnum.IsDefined<ResSplitByMonthAttribute>() == true)
            {
                fileName = Path.Combine(DateTime.Now.ToString("yyyy_MM"), fileName);
            }

            string fullFileName = Path.Combine(ResManage.RootFullPath, resType.ToString(), fileName);
            string fullPath = Path.GetDirectoryName(fullFileName);
            Directory.CreateDirectory(fullPath);
            file.SaveAs(fullFileName);

            return ResManage.Parse(fileName, resType);
        }
    }
}