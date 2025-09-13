using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SVGControl
{
    [TypeConverter(typeof(SvgResourceConverter))]
    public interface ISvgResource
    {
        string Name { get; }
        byte[] Data { get; }
    }

    public class SvgResource : ISvgResource
    {
        public SvgResource() { }

        public SvgResource(string name, byte[] data)
        {
            Name = name;
            Data = data;
        }

        public string Name { get; set; }
        public byte[] Data { get; set; }
    }
}
