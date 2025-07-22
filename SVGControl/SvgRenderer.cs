using Svg;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SVGControl
{
    internal class SvgRenderer : INotifyPropertyChanged
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public SvgRenderer(byte[] doc, Size size, AutoSize autoSize)
        {
            _doc = GetSvgDocument(doc);
            _original = _doc.Draw().Size;
            _margin = new Padding(0);
            Size = CalcInnerSize(size, _margin);
            _autoSize = autoSize;
        }

        public SvgRenderer(byte[] doc, Size size, Padding margin, AutoSize autoSize)
        {
            _doc = GetSvgDocument(doc);
            _original = _doc.Draw().Size;
            _margin = margin;
            Size = CalcInnerSize(size, _margin);
            _autoSize = autoSize;
        }

        public SvgRenderer(SvgDocument doc, Size size, AutoSize autoSize)
        {
            _doc = doc;
            _original = _doc.Draw().Size;
            _margin = new Padding(0);
            Size = CalcInnerSize(size, _margin);
            _autoSize = autoSize;
        }
        
        public SvgRenderer(SvgDocument doc, Size size, Padding margin, AutoSize autoSize)
        {
            _doc = doc;
            _original = _doc.Draw().Size;
            _margin = margin;
            Size = CalcInnerSize(size, _margin);
            _autoSize = autoSize;
        }

        public SvgRenderer(Size outer, Padding margin, AutoSize autoSize)
        {
            _outer = outer;
            Margin = margin;
            AutoSize = autoSize;
            Size = CalcInnerSize(outer, margin);
            //logger.Debug("SvgRenderer Initialized");
        }

        private Size _outer;
        private Size _original; 
        private Padding _margin;
        private SvgDocument _doc;
        private AutoSize _autoSize;
        private Size _size;
                
        [NotifyParentProperty(true)]
        internal Size Outer
        {
            get { return _outer; }
            set
            {
                _outer = value;
                Size = CalcInnerSize(Outer, _margin);
                NotifyPropertyChanged("Outer");
            }
        }

        [NotifyParentProperty(true)]
        public Size Size { get => _size; set => _size = value; }

        [NotifyParentProperty(true)]
        public Padding Margin
        {
            get { return _margin; }
            set
            {
                _margin = value;
                Size = CalcInnerSize(Outer, _margin);
                NotifyPropertyChanged("Margin");
            }
        }

        [NotifyParentProperty(true)]
        [DefaultValue(AutoSize.MaintainAspectRatio)]
        public AutoSize AutoSize { get => _autoSize; set => _autoSize = value; }

        [NotifyParentProperty(true)]
        public SvgDocument Document 
        { 
            get => _doc;
            set 
            { 
                _doc = value;
                if (value != null) { _original = _doc.Draw().Size; }
                NotifyPropertyChanged();
            }
        }
        
        private Size CalcInnerSize(Size outer, Padding margin)
        {
            var innerWidth = outer.Width - margin.Left - margin.Right;
            var innerHeight = outer.Height - margin.Top - margin.Bottom;
            return new Size(innerWidth, innerHeight);
        }

        public Bitmap Render()
        {
            if (_doc == null)
            {
                return null;
            }
            else if ((AutoSize == AutoSize.Disabled) || (Size == null) || (Size.Height == 0) || (Size.Width == 0))
            {
                return _doc.Draw();
            }
            else if (AutoSize == AutoSize.AllowStretching)
            {
                _doc.Width = Size.Width;
                _doc.Height = Size.Height;
                return _doc.Draw();
            }
            else if (AutoSize == AutoSize.MaintainAspectRatio)
            {
                var targetAdjusted = AdjustSizeProportionately(_original, Size);
                _doc.Width = targetAdjusted.Width;
                _doc.Height = targetAdjusted.Height;
                //AddMargins(targetAdjusted.Width, targetAdjusted.Height);
                return _doc.Draw();
            }
            else
            { return null; }
        }

        private void AddMargins(int widthCurrent, int heightCurrent)
        {
            var group = new SvgGroup();
            _doc.Children.Add(group);
            group.Children.Add(new SvgRectangle
            {
                X = -_margin.Left,
                Y = -_margin.Top,
                Width = widthCurrent + Margin.Left + Margin.Right,
                Height = heightCurrent + Margin.Top + Margin.Bottom,
                Stroke = new SvgColourServer(Color.Transparent),
                Fill = new SvgColourServer(Color.Transparent)
            });
        }

        private Size AdjustSizeProportionately(Size proportions, Size targetSize)
        {
            if ((targetSize.Height > 0) && (targetSize.Width > 0) && ((proportions.Height != targetSize.Height) || (proportions.Width != targetSize.Width)))
            {
                int widthAspect = (int)(targetSize.Height * proportions.Width / (double)proportions.Height);
                if (widthAspect < targetSize.Width)
                {
                    return new Size(widthAspect, targetSize.Height);

                }
                else
                {
                    int heightAspect = (int)(targetSize.Width * proportions.Height / (double)proportions.Width);
                    return new Size(targetSize.Width, heightAspect);
                }
            }
            return proportions;
        }

        public static SvgDocument GetSvgDocument(byte[] file)
        {
            Stream stream = new MemoryStream(file);
            try
            {
                return SvgDocument.Open<SvgDocument>(stream);
            }
            catch (Exception)
            {
                return null;
            }
        }

        #region EventHandlers
        
        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        
        #endregion
    }
}
