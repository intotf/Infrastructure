using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.IO;
using Aspose.Cells;
using System.Drawing;
using Infrastructure.Reflection;
using Infrastructure.Attributes;

namespace Infrastructure.Utility
{

    /// <summary>
    /// 定义导入的实体属性 特性
    /// </summary>
    public class CellColumnAttribute : Attribute
    {
        /// <summary>
        /// 对应 Excle 第一行名字
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="name">Excle 第一行表头</param>
        public CellColumnAttribute(string name)
        {
            this.Name = name;
        }
    }

    /// <summary>
    /// 读取 Excle 文件做导入
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class Cells<T> where T : new()
    {
        /// <summary>
        /// 导出字段
        /// </summary>
        private List<IField> fields = new List<IField>();

        /// <summary>
        /// 获取数据模型
        /// </summary>
        public IEnumerable<T> Models { get; private set; }

        /// <summary>
        /// 数据转换为Excel
        /// </summary>
        /// <param name="models">数据</param>
        public Cells(IEnumerable<T> models)
        {
            this.Models = models;
        }

        /// <summary>
        /// 导入文件
        /// </summary>
        /// <param name="fileName">文件绝对路径</param>
        public Cells(string fileName)
        {
            this.Models = this.ReadCells(fileName);
        }

        /// <summary>
        /// 读取Excle 第一个Sheet 对象集合
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private IEnumerable<T> ReadCells(string fileName)
        {
            var cells = new Workbook(fileName).Worksheets[0].Cells;
            var firstRow = cells.Rows.First();
            var head = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i <= cells.MaxDataColumn; i++)
            {
                var key = firstRow.GetCellByIndex(i).Value;
                head.Add(key.ToString(), i);
            }

            var properties = Property.GetProperties(typeof(T));

            foreach (var row in cells.Rows.Skip(1))
            {
                var model = new T();
                foreach (var p in properties)
                {
                    var column = Attribute.GetCustomAttribute(p.Info, typeof(CellColumnAttribute)) as CellColumnAttribute;
                    var pName = column == null ? p.Name : column.Name;
                    //判断字段是否存在
                    if (head.ContainsKey(pName))
                    {
                        var index = head[pName];
                        var value = row.GetCellOrNull(index) == null ? null : row.GetCellOrNull(index).Value;
                        var valueCast = Converter.Cast(value, p.Info.PropertyType);
                        p.SetValue(model, valueCast);
                    }
                }
                yield return model;
            }
        }

        /// <summary>
        /// 添加字段
        /// </summary>
        /// <typeparam name="TField">字段类型</typeparam>   
        /// <param name="name">字段名</param>
        /// <param name="value">字段值</param>
        public Cells<T> AddField<TField>(string name, Func<T, TField> value)
        {
            var field = new Field<TField>(name, value);
            this.fields.Add(field);
            return this;
        }

        /// <summary>
        /// 标题头简单样式
        /// </summary>
        /// <returns></returns>
        private Style GetHeaderStyle()
        {
            var style = new Style();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.Font.Name = "Arial";
            style.Font.Size = 11;
            style.IsTextWrapped = false;
            style.Font.IsBold = true;
            return style;
        }

        /// <summary>
        /// 保存到文件
        /// </summary>
        /// <param name="stream">流</param>
        public void Save(Stream stream)
        {
            var xls = new Aspose.Cells.Workbook();

            var sheet = xls.Worksheets[0];
            var style = this.GetHeaderStyle();

            for (var i = 0; i < this.fields.Count; i++)
            {
                var header = sheet.Cells[0, i];
                header.SetStyle(style, true);
                header.PutValue(this.fields[i].Name);
            }

            var row = 0;
            foreach (var model in this.Models)
            {
                row = row + 1;
                for (var i = 0; i < this.fields.Count; i++)
                {
                    var fieldValue = this.fields[i].GetValue(model);
                    sheet.Cells[row, i].PutValue(fieldValue);
                }
            }

            sheet.AutoFitColumns();
            xls.Save(stream, Aspose.Cells.SaveFormat.Excel97To2003);
        }

        /// <summary>
        /// 字段接口
        /// </summary>
        private interface IField
        {
            string Name { get; }
            object GetValue(T model);
        }

        /// <summary>
        /// 表示字段信息
        /// </summary>
        /// <typeparam name="TKey"></typeparam>
        private class Field<TKey> : IField
        {
            private Func<T, TKey> valueFunc;

            public string Name { get; private set; }

            public Field(string name, Func<T, TKey> value)
            {
                this.Name = name;
                this.valueFunc = value;
            }

            public object GetValue(T model)
            {
                return this.valueFunc(model);
            }
        }
    }
}
