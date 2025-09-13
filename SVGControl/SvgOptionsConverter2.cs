using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SVGControl
{
    public class SvgOptionsConverter : ExpandableObjectConverter
    {
        public override object ConvertTo(
            ITypeDescriptorContext context,
            CultureInfo culture,
            object value,
            Type destinationType)
        {
            if (destinationType == typeof(string))
            {
                SvgImageSelector image = value as SvgImageSelector;
                if (image != null)
                {
                    if (image.ResourceName != null)
                    {
                        string resourceName = image.ResourceName.Name;
                        string autoSizeCode;
                        switch (image.AutoSize)
                        {
                            case AutoSize.Disabled: autoSizeCode = "[Static]"; break;
                            case AutoSize.MaintainAspectRatio: autoSizeCode = "[Proportional]"; break;
                            case AutoSize.AllowStretching: autoSizeCode = "[Stretchable]"; break;
                            default: autoSizeCode = "[]"; break;
                        }

                        return $"{resourceName} {autoSizeCode}";
                    }

                    //if (image.AboluteImagePath != null)
                    //{
                    //    string filename = Path.GetFileName(image.AboluteImagePath);
                    //    string autoSizeCode;
                    //    switch (image.AutoSize)
                    //    {
                    //        case AutoSize.Disabled: autoSizeCode = "[Static]"; break;
                    //        case AutoSize.MaintainAspectRatio: autoSizeCode = "[Proportional]"; break;
                    //        case AutoSize.AllowStretching: autoSizeCode = "[Stretchable]"; break;
                    //        default: autoSizeCode = "[]"; break;
                    //    }

                    //    return $"{filename} {autoSizeCode}";
                    //}
                    else { return "(none)"; }

                }
                return "";
            }

            return base.ConvertTo(
                context,
                culture,
                value,
                destinationType);
        }
    }

}
