using Svg;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SVGControl
{
    internal class SVGParser
    {
        /// <summary>
        /// The maximum image _size supported.
        /// </summary>
        public Size TargetSize { get; set; }

        /// <summary>
        /// Converts an SVG file to a Bitmap image.
        /// </summary>
        /// <param name="filePath">The full path of the SVG image.</param>
        /// <returns>Returns the converted Bitmap image.</returns>
        public Bitmap GetBitmapFromSVG(string filePath)
        {
            SvgDocument document = GetSvgDocument(filePath);
            if (TargetSize != null)
                AdjustSize(document);
            Bitmap bmp = document.Draw();
            return bmp;
        }

        public Bitmap GetBitmapFromSVG(string filePath, Size size)
        {
            SvgDocument document = GetSvgDocument(filePath);
            TargetSize = size;
            AdjustSize(document);
            Bitmap bmp = document.Draw();
            return bmp;
        }

        public Bitmap GetBitmapFromSVG(byte[] file, Size size)
        {
            SvgDocument document = GetSvgDocument(file);
            TargetSize = size;
            AdjustSize(document);
            Bitmap bmp = document.Draw();
            return bmp;
        }

        public Bitmap GetBitmapFromSVG(string filePath, int Height, int Width)
        {
            SvgDocument document = GetSvgDocument(filePath);
            TargetSize = new Size(Height, Width);
            AdjustSize(document);
            Bitmap bmp = document.Draw();
            return bmp;
        }

        /// <summary>
        /// Gets a SvgDocument for manipulation using the path provided.
        /// </summary>
        /// <param name="filePath">The path of the Bitmap image.</param>
        /// <returns>Returns the SVG Document.</returns>
        public SvgDocument GetSvgDocument(string filePath)
        {
            SvgDocument document = SvgDocument.Open(filePath);
            return document;
        }

        public SvgDocument GetSvgDocument(byte[] file)
        {
            Stream stream = new MemoryStream(file);
            SvgDocument document = SvgDocument.Open<SvgDocument>(stream);
            return document;
        }

        /// <summary>
        /// Makes sure that the image does not exceed the maximum _size, while preserving aspect ratio.
        /// </summary>
        /// <param name="document">The SVG document to resize.</param>
        /// <returns>Returns a resized or the original document depending on the document.</returns>
        private SvgDocument AdjustSize(SvgDocument document)
        {
            if ((TargetSize.Height > 0) && (TargetSize.Width > 0) && ((document.Height != TargetSize.Height)||(document.Width != TargetSize.Width))) 
            { 
                int widthAspect = (int)(TargetSize.Height * document.Width/(double)document.Height);
                if (widthAspect < TargetSize.Width) 
                {
                    document.Height = TargetSize.Height;
                    document.Width = widthAspect;
                }
                else
                {
                    document.Height = (int)(TargetSize.Width * document.Height/(double)document.Width);
                    document.Width = TargetSize.Width;
                }
            }
            return document;
        }
    }
}
