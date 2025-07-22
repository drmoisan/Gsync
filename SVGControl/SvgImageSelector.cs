using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using Svg;
using System.Runtime.CompilerServices;


namespace SVGControl
{
    public enum AutoSize
    {
        Disabled = 0,
        MaintainAspectRatio = 1,
        AllowStretching = 2
    }

    [TypeConverter(typeof(SvgOptionsConverter))]
    public class SvgImageSelector : INotifyPropertyChanged
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public SvgImageSelector(Size outer, Padding margin, AutoSize autoSize)
        {
            ResourceNames = new List<ISvgResource>();
            _renderer = new SvgRenderer(outer, margin, autoSize);
            _renderer.PropertyChanged += Renderer_PropertyChanged;
            //logger.Debug("SvgImageSelector Initialized");
        }

        public SvgImageSelector(Size outer, Padding margin, AutoSize autoSize, bool useDefaultImage)
        {
            ResourceNames = new List<ISvgResource>();
            _useDefaultImage = useDefaultImage;
            if (useDefaultImage)
            {
                _renderer = new SvgRenderer(Defaults.GetDefault.SvgImage,
                                            outer,
                                            margin,
                                            autoSize);
            }
            else { _renderer = new SvgRenderer(outer, margin, autoSize); }
            _renderer.PropertyChanged += Renderer_PropertyChanged;
            //logger.Debug("SvgImageSelector Initialized");
        }

        //private SvgDocument _doc;
        private string _relativeImagePath;
        private string _absoluteImagePath;
        private bool _saveRendering = false;
        private bool _useDefaultImage = false;
        private SvgRenderer _renderer;

        internal String AboluteImagePath
        {
            get { return _absoluteImagePath; }
        }

        #region Public Properties

        [NotifyParentProperty(true)]
        [Editor(typeof(SvgFileNameEditor), typeof(UITypeEditor))]
        public string ImagePath
        {
            get
            {
                if (_absoluteImagePath == null)
                {
                    return "(none)";
                }
                else
                {
                    return _relativeImagePath;
                }

            }
            set
            {
                //if (_relativeImagePath != value)
                //{
                //    if ((value == "")|(value == "(none)"))
                //    {
                //        _relativeImagePath = value;
                //        if (_useDefaultImage) { SetDefaultImage(); }
                //        else { _renderer.Document = null; }
                //    }
                //    else
                //    {
                //        string valueAbs = value.AbsoluteFromURI(GetAnchorPath());
                //        _renderer.Document = SvgDocument.Open(valueAbs);
                //        _absoluteImagePath = valueAbs;
                //        _relativeImagePath = valueAbs.GetRelativeURI(GetAnchorPath());
                //    }
                //    NotifyPropertyChanged("ImagePath");
                //}
            }
        }

        private ISvgResource _svgResource = null;

        [Editor(typeof(DropDownEditor), typeof(UITypeEditor))]
        public ISvgResource ResourceName
        {
            get => _svgResource ;
            
            set
            {
                if (_svgResource != value)
                {
                    if (value is null || value.Name == "")
                    {
                        _svgResource = value;
                        if (_useDefaultImage) { SetDefaultImage(); }
                        else { _renderer.Document = null; }
                    }
                    else
                    {
                        _svgResource = value;
                        _renderer.Document = SvgRenderer.GetSvgDocument(value.Data);
                    }
                    NotifyPropertyChanged("ResourceName");
                }
                //_svgResource = value;
            }
        }

        //private List<string> _resourceNames;

        [Browsable(false)]
        public List<ISvgResource> ResourceNames { get; private set; }

        public void AddResource(ISvgResource resource)
        {
            if (!ResourceNames.Contains(resource))
            ResourceNames.Add(resource);
        }

        [DefaultValue(AutoSize.MaintainAspectRatio)]
        public AutoSize AutoSize { get => _renderer.AutoSize; set => _renderer.AutoSize = value; }
        
        public Size Size { get => _renderer.Size; set => _renderer.Size = value; }
        
        public Padding Margin { get => _renderer.Margin; set => _renderer.Margin = value; }
        
        public Bitmap Render() => _renderer.Render();

        public bool UseDefaultImage
        {
            get => _useDefaultImage;
            set
            {
                _useDefaultImage = value;
                if ((_relativeImagePath == "") | (_relativeImagePath == "(none)"))
                {
                    if (_useDefaultImage) { SetDefaultImage(); }
                    else { _renderer.Document = null; }
                }
                    
            }
        }
        
        public bool SaveRendering
        {
            get
            {
                return _saveRendering;
            }
            set
            {
                if ((value == true) && (_relativeImagePath != "") && (_renderer.Document != null))
                {
                    // Launch file save dialog with appropriate filters
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Png Image|*.png|JPeg Image|*.jpg|Bitmap Image|*.bmp|Gif Image|*.gif";
                    saveFileDialog1.Title = "Save rendered Image File";
                    saveFileDialog1.InitialDirectory = Path.GetFullPath(_relativeImagePath);
                    saveFileDialog1.FileName = Path.GetFileNameWithoutExtension(_relativeImagePath);

                    saveFileDialog1.ShowDialog();

                    // If the file name is not an empty string open it for saving.
                    if (saveFileDialog1.FileName != "")
                    {
                        // Saves the Image via a FileStream created by the OpenFile method.
                        using FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
                        {
                            Image image = Render();
                            // Saves the Image in the appropriate ImageFormat based upon the
                            // File type selected in the dialog box.
                            // NOTE that the FilterIndex property is one-based.
                            switch (saveFileDialog1.FilterIndex)
                            {
                                case 1:
                                    image.Save(fs, System.Drawing.Imaging.ImageFormat.Png);
                                    break;
                                case 2:
                                    image.Save(fs, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    break;
                                case 3:
                                    image.Save(fs, System.Drawing.Imaging.ImageFormat.Bmp);
                                    break;
                                case 4:
                                    image.Save(fs, System.Drawing.Imaging.ImageFormat.Gif);
                                    break;
                            }
                        } // end using FileStream fs
                    }

                }
                else if ((value == true) && (_relativeImagePath == ""))
                {
                    MessageBox.Show("Image path must have a value to save the rendering");

                }
                else if (_renderer.Document == null)
                {
                    // MessageBox.Show("Image path does not refer to a valid SVG document");

                }
                _saveRendering = false;

            }
        }

        #endregion

        #region Internal and Private Functions

        internal Size Outer { get => _renderer.Outer; set => _renderer.Outer = value; }

        private string GetAnchorPath()
        {
            //string workingDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            //Debug.WriteLine("System.AppDomain.CurrentDomain.BaseDirectory is " + workingDirectory);
            //var assmbly = System.Reflection.Assembly.GetExecutingAssembly();
            //var location = assmbly.Location;
            
            //var resrcs = assmbly.GetManifestResourceNames();
            
            
            string workingDirectory = Environment.CurrentDirectory;
            List<string> directories = new List<string>(workingDirectory.Split(Path.DirectorySeparatorChar));
            if ((directories.Count > 2) && (directories[directories.Count - 2] == "bin"))
            {
                // Backwards traverse 2 levels
                workingDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            }
            if (workingDirectory[workingDirectory.Length - 1] != Path.DirectorySeparatorChar)
            {
                workingDirectory += Path.DirectorySeparatorChar;
            }
            return workingDirectory;
        }
                
        internal void SetDefaultImage()
        {
            _renderer.Document = SvgRenderer.GetSvgDocument(Defaults.GetDefault.SvgImage);
        }

        #endregion

        #region EventHandlers

        public event PropertyChangedEventHandler PropertyChanged;

        private void Renderer_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            PropertyChanged?.Invoke(this, e);
        }

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion


    }


}

namespace SVGControl.Defaults
{
    public static class GetDefault
    {
        public static byte[] SvgImage 
        { 
            get 
            {
                string svgXML = 
@"<svg xmlns=""http://www.w3.org/2000/svg"" viewBox=""0 0 16 16"">
  <defs>
    <style>.canvas{fill: none; opacity: 0;}.light-defaultgrey-10{fill: #212121; opacity: 0.1;}.light-defaultgrey{fill: #212121; opacity: 1;}.light-yellow{fill: #996f00; opacity: 1;}.light-blue{fill: #005dba; opacity: 1;}</style>
  </defs>
  <title>IconLightImage</title>
  <g id=""canvas"" class=""canvas"">
    <path class=""canvas"" d=""M16,16H0V0H16Z"" />
  </g>
  <g id=""level-1"">
    <path class=""light-defaultgrey-10"" d=""M14.5,2.5v12H1.5V2.5Z"" />
    <path class=""light-defaultgrey"" d=""M14.5,2H1.5L1,2.5v12l.5.5h13l.5-.5V2.5ZM14,14H2V3H14Z"" />
    <path class=""light-yellow"" d=""M12,5.5A1.5,1.5,0,1,1,10.5,4,1.5,1.5,0,0,1,12,5.5Z"" />
    <path class=""light-blue"" d=""M14,11.09V12.5l-2.819-2.82L8.988,11.877H8.281L4.814,8.41,2,11.225V9.811L4.461,7.35h.707l3.466,3.466,2.193-2.193h.707Z"" />
  </g>
</svg>";
                return Encoding.ASCII.GetBytes(svgXML);
            } 
        }
        
    }
}
