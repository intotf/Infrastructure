using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;

namespace Infrastructure.Resource
{
    /// <summary>
    /// 图片资源类
    /// 提供图片缩放、添加水印功能   
    /// </summary>
    public sealed class ImageRes : IDisposable
    {
        /// <summary>
        /// 从图片文件名获取图片资源
        /// </summary>
        /// <param name="fileFullName">图片文件名</param>
        /// <exception cref="FileNotFoundException"></exception>
        /// <returns></returns>
        public static ImageRes Parse(string fileFullName)
        {
            if (File.Exists(fileFullName) == false)
            {
                throw new FileNotFoundException();
            }
            return new ImageRes(fileFullName);
        }

        /// <summary>
        /// 缩放模式
        /// </summary>
        public enum ScaleMode
        {
            /// <summary>
            /// 使X轴铺满，Y轴保持不变
            /// </summary>
            Fill_X,
            /// <summary>
            /// 使Y轴铺满，X轴保持不变
            /// </summary>
            Fill_Y,
            /// <summary>
            /// 使XY都轴铺满
            /// </summary>
            Fill_XY,
            /// <summary>
            /// 使X轴铺满，Y轴跟着比例缩放
            /// </summary>
            Fill_X_WithY,
            /// <summary>
            /// 使Y轴铺满，X轴跟着比例缩放
            /// </summary>
            Fill_Y_WithX,
            /// <summary>
            /// 使X或Y轴尽量铺满，两轴缩放比例一样
            /// </summary>
            Fill_XY_Auto,

            /// <summary>
            /// 使X和Y以最小边等比铺满，两轴缩放比较一样
            /// </summary>
            Fill_XY_Center,
        }

        /// <summary>
        /// 缩放选项
        /// </summary>
        public enum ScaleOption
        {
            /// <summary>
            /// 总是缩放
            /// </summary>
            AllWay,
            /// <summary>
            /// 当图片尺寸比目标小时不缩放
            /// </summary>
            NoSacleWhenSmall
        }

        /// <summary>
        /// 水印位置
        /// </summary>
        public enum MarkPosition
        {
            /// <summary>
            /// 左上
            /// </summary>
            TopLeft,
            /// <summary>
            /// 右上
            /// </summary>
            TopRiht,
            /// <summary>
            /// 左下
            /// </summary>
            BottomLeft,
            /// <summary>
            /// 右下
            /// </summary>
            BottomRight,
            /// <summary>
            /// 居中
            /// </summary>
            Center
        }


        /// <summary>
        /// 图像
        /// </summary>
        private Bitmap bitmap;

        /// <summary>
        /// 获取图像的文件名
        /// </summary>
        public string ImageFileName { get; private set; }

        /// <summary>
        /// 获取图像的高度
        /// </summary>
        public int Height
        {
            get
            {
                return this.bitmap.Height;
            }
        }

        /// <summary>
        /// 获取图像的宽度
        /// </summary>
        public int Width
        {
            get
            {
                return this.bitmap.Width;
            }
        }

        /// <summary>
        /// 构造器
        /// </summary>
        /// <param name="imageFullName">图片文件名</param>
        private ImageRes(string imageFullName)
        {
            this.ImageFileName = imageFullName;
            this.bitmap = new Bitmap(new MemoryStream(File.ReadAllBytes(imageFullName)));

        }

        /// <summary>
        /// 以某种比例缩放图片
        /// </summary>
        /// <param name="newSize">新的尺寸</param>
        /// <param name="scaleMode">缩放模式</param>
        /// <param name="scaleOpiton">缩放选项</param>
        public void SacleTo(Size newSize, ScaleMode scaleMode, ScaleOption scaleOpiton = ScaleOption.NoSacleWhenSmall)
        {
            Size size = this.bitmap.Size;

            if (scaleOpiton == ScaleOption.NoSacleWhenSmall)
            {
                if (newSize.Height >= size.Height && newSize.Width >= size.Width)
                {
                    return;
                }
            }

            newSize = this.GetImageNewSize(size, newSize, scaleMode);
            var img = new Bitmap(this.bitmap, newSize);
            this.bitmap.Dispose();
            this.bitmap = img;
        }

        /// <summary>
        /// 居中裁剪
        /// </summary>
        public void CutForSquare(Size size)
        {
            this.SacleTo(size, ScaleMode.Fill_XY_Center);

            var initImage = this.bitmap;
            //原始图片的宽、高
            int initWidth = initImage.Width;
            int initHeight = initImage.Height;

            //对象实例化
            var pickedImage = new System.Drawing.Bitmap(initHeight, initHeight);
            var pickedG = System.Drawing.Graphics.FromImage(pickedImage);

            //设置质量
            pickedG.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            pickedG.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

            if (initWidth > initHeight)
            {
                var x = (initWidth - size.Width) / 2;
                //定位
                Rectangle fromR = new Rectangle(x, 0, size.Width, size.Height);
                Rectangle toR = new Rectangle(0, 0, initHeight, initHeight);
                //画图
                pickedG.DrawImage(initImage, toR, fromR, GraphicsUnit.Pixel);
            }
            else
            {
                var y = (initHeight - size.Height) / 2;
                //定位
                Rectangle fromR = new Rectangle(0, y, size.Width, size.Height);
                Rectangle toR = new Rectangle(0, 0, initHeight, initHeight);
                pickedG.DrawImage(initImage, toR, fromR, GraphicsUnit.Pixel);
            }

            initImage = (System.Drawing.Bitmap)pickedImage.Clone();
            var img = new Bitmap(initImage, size);
            this.bitmap = img;

            //释放截图资源
            pickedG.Dispose();
            pickedImage.Dispose();
        }

        /// <summary>
        /// 获取图像的新大小
        /// </summary>
        /// <param name="imageSize">图像大小</param>
        /// <param name="newSize">预设的新大小</param>
        /// <param name="scaleMode">缩放模式</param>
        /// <returns></returns>
        private Size GetImageNewSize(Size imageSize, Size newSize, ScaleMode scaleMode)
        {
            switch (scaleMode)
            {
                case ScaleMode.Fill_X:
                    newSize.Height = imageSize.Height;
                    break;

                case ScaleMode.Fill_X_WithY:
                    decimal scale = (decimal)imageSize.Width / (decimal)newSize.Width;
                    decimal heigth = (decimal)imageSize.Height / scale;
                    newSize.Height = (int)heigth;
                    break;

                case ScaleMode.Fill_Y:
                    newSize.Width = imageSize.Width;
                    break;

                case ScaleMode.Fill_Y_WithX:
                    scale = (decimal)imageSize.Height / (decimal)newSize.Height;
                    decimal width = (decimal)imageSize.Width / scale;
                    newSize.Width = (int)width;
                    break;

                case ScaleMode.Fill_XY:
                    break;

                case ScaleMode.Fill_XY_Auto:
                    decimal sacleX = (decimal)imageSize.Width / (decimal)newSize.Width;
                    decimal scaleY = (decimal)imageSize.Height / (decimal)newSize.Height;
                    scale = Math.Max(sacleX, scaleY);
                    newSize.Width = (int)(imageSize.Width / scale);
                    newSize.Height = (int)(imageSize.Height / scale);
                    break;

                case ScaleMode.Fill_XY_Center:

                    if (imageSize.Width > imageSize.Height)
                    {
                        decimal x = (decimal)imageSize.Height / (decimal)newSize.Height;

                        newSize.Width = (int)(imageSize.Width / x);
                        newSize.Height = (int)(imageSize.Height / x);
                    }
                    else
                    {
                        decimal y = (decimal)imageSize.Width / (decimal)newSize.Width;

                        newSize.Width = (int)(imageSize.Width / y);
                        newSize.Height = (int)(imageSize.Height / y);
                    }
                    break;

            }
            return newSize;
        }

        /// <summary>
        /// 添加水印文字
        /// </summary>
        /// <param name="markString">水印文字</param>
        /// <param name="font">字段</param>
        /// <param name="markPosition">位置</param>
        /// <param name="margin">边距</param>
        public void DrawMarkString(string markString, Font font, MarkPosition markPosition, int margin = 10)
        {
            using (Graphics g = Graphics.FromImage(this.bitmap))
            {
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                var sizeF = g.MeasureString(markString, font);
                // 文本区域大小
                var objectSize = new Size((int)sizeF.Width, (int)sizeF.Height);
                // 绘制水印的左上角点
                var point = this.GetObjectDrawPoint(objectSize, markPosition, margin);

                #region 计算文本的颜色
                var centerP = new Point(point.X + objectSize.Width / 2, point.Y - objectSize.Height / 2);
                var backColor = Color.White;
                var rect = new Rectangle(Point.Empty, this.bitmap.Size);
                if (rect.Contains(centerP))
                {
                    backColor = this.bitmap.GetPixel(centerP.X, centerP.Y);
                }
                var foreColor = Color.FromArgb(255 - backColor.R, 255 - backColor.G, 255 - backColor.B);
                #endregion

                g.DrawString(markString, font, new SolidBrush(foreColor), point);
            }
        }


        /// <summary>
        /// 绘制图片水印
        /// </summary>
        /// <param name="markImageFullName">水印图片文件名</param>
        /// <param name="markPosition">水印位置</param>
        /// <param name="margin">外边距</param>
        public void DrawMarkImage(string markImageFullName, MarkPosition markPosition, int margin = 10)
        {
            using (Graphics g = Graphics.FromImage(this.bitmap))
            {
                using (var image = new Bitmap(new MemoryStream(File.ReadAllBytes(markImageFullName))))
                {
                    // 绘制水印的左上角点
                    var point = this.GetObjectDrawPoint(image.Size, markPosition, margin);
                    g.DrawImage(image, point);
                }
            }
        }


        /// <summary>
        /// 获取绘制物体的八绘制左上角点坐标
        /// </summary>
        /// <param name="objectSize">物体大小</param>
        /// <param name="markPosition">水印位置</param>
        /// <param name="margin">外边距</param>
        /// <returns></returns>
        private Point GetObjectDrawPoint(Size objectSize, MarkPosition markPosition, int margin)
        {
            Point p = Point.Empty;
            switch (markPosition)
            {
                case MarkPosition.BottomLeft:
                    p.X = margin;
                    p.Y = this.bitmap.Height - objectSize.Height - margin;
                    break;

                case MarkPosition.BottomRight:
                    p.X = this.bitmap.Width - objectSize.Width - margin;
                    p.Y = this.bitmap.Height - objectSize.Height - margin;
                    break;

                case MarkPosition.TopLeft:
                    p.X = margin;
                    p.Y = margin;
                    break;

                case MarkPosition.TopRiht:
                    p.X = this.bitmap.Width - objectSize.Width - margin;
                    p.Y = margin;
                    break;

                case MarkPosition.Center:
                    p.X = this.bitmap.Width / 2 - objectSize.Width / 2;
                    p.Y = this.bitmap.Height / 2 - objectSize.Height / 2;
                    break;
            }
            return p;
        }


        /// <summary>
        /// 获取质量参数
        /// </summary>
        /// <param name="qaulity">质量</param>
        /// <returns></returns>
        private EncoderParameters GetQaulityEncoderParameter(int qaulity)
        {
            var encoderParameter = new EncoderParameter(Encoder.Quality, qaulity);
            return new EncoderParameters { Param = new EncoderParameter[] { encoderParameter } };
        }

        /// <summary>
        /// 获取编码信息
        /// </summary>
        /// <returns></returns>
        private ImageCodecInfo GetCodecInfo()
        {
            var ext = Path.GetExtension(this.ImageFileName).ToUpper();
            return ImageCodecInfo.GetImageEncoders().FirstOrDefault(item => item.FilenameExtension.ToUpper().Contains(ext));
        }

        /// <summary>
        /// 保存并替换原始图像文件        
        /// </summary>
        public void Save()
        {
            this.bitmap.Save(this.ImageFileName);
        }

        /// <summary>
        /// 保存并替换原始图像文件
        /// <param name="qaulity">质量</param>
        /// </summary>
        public void Save(int qaulity)
        {
            var codec = this.GetCodecInfo();
            var parameter = this.GetQaulityEncoderParameter(qaulity);
            this.bitmap.Save(this.ImageFileName, codec, parameter);
        }

        /// <summary>
        /// 保存图片到流
        /// </summary>
        /// <param name="stream">数据流</param>
        public void Save(Stream stream)
        {
            this.bitmap.Save(stream, this.bitmap.RawFormat);
        }


        /// <summary>
        /// 保存图片到流
        /// </summary>
        /// <param name="stream">数据流</param>
        /// <param name="qulity">质量</param>
        public void Save(Stream stream, int qaulity)
        {
            var codec = this.GetCodecInfo();
            var parameter = this.GetQaulityEncoderParameter(qaulity);
            this.bitmap.Save(stream, codec, parameter);
        }


        /// <summary>
        /// 图像名
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return this.ImageFileName;
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            if (this.bitmap != null)
            {
                this.bitmap.Dispose();
                this.bitmap = null;
            }
        }
    }
}