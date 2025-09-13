using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.Design;
using System.Windows.Forms;
using System.ComponentModel.Design;
using System.Reflection;
using System.Globalization;
using System.Resources;


namespace SVGControl
{
    // this defines a custom UI type editor to display a list of possible benchmarks
    // used by the property grid to display item in edit mode
    public class DropDownEditor : UITypeEditor
    {
        private IWindowsFormsEditorService _editorService;

        public override UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext context)
        {
            // drop down mode (we'll host a listbox in the drop down)
            return UITypeEditorEditStyle.DropDown;
        }

        public override object EditValue(ITypeDescriptorContext context, IServiceProvider provider, object value)
        {
            IDesignerHost host = provider.GetService(typeof(IDesignerHost)) as IDesignerHost;
            string typName = host.RootComponentClassName;
            Type typ = host.GetType(typName);
            Assembly asm = null;
            if (typ == null)
            {
                MessageBox.Show("Please build project before attempting to set this property");
                return base.EditValue(context, provider, value);
            }
            else
            {
                asm = typ.Assembly;
            }
            var assemblyName = asm.GetName().Name;
            var rm = new ResourceManager($"{assemblyName}.Properties.Resources", asm);
            var rset = rm.GetResourceSet(CultureInfo.CurrentUICulture, true, true);


            var entries2 = rset.Cast<System.Collections.DictionaryEntry>()
                              .OrderBy(x => x.Key)
                              .Select(x => new KeyValuePair<string,object>((string)x.Key, x.Value))
                              .ToList();

            var entries = rset.Cast<System.Collections.DictionaryEntry>()
                              .Where(x => x.Value is byte[])
                              .OrderBy(x => x.Key)
                              .Select(x => new SvgResource((string)x.Key, (byte[])x.Value))
                              .ToList();

            // get the analytic object from context
            // this is how we get the list of possible benchmarks
            SvgImageSelector imageSelector = (SvgImageSelector)context.Instance;

            //foreach (var entry in entries)
            //{
            //    imageSelector.AddResource(entry);
            //}
            
            //var names = rset.Cast<System.Collections.DictionaryEntry>()
            //                .Where(x => x.Value is byte[])
            //                .Select(x => GetStringForValue(x.Key))
            //                .ToList();

            _editorService = (IWindowsFormsEditorService)provider.GetService(typeof(IWindowsFormsEditorService));

            // use a list box
            ListBox lb = new ListBox();
            lb.SelectionMode = SelectionMode.One;
            //lb.SelectedValueChanged += OnListBoxSelectedValueChanged;

            // use the IBenchmark.Name property for list box display
            lb.DisplayMember = "Name";


            //foreach (string resourceName in names)
            //foreach (string resourceName in imageSelector.ResourceNames)
            lb.Items.Clear();
            //foreach (ISvgResource resourceName in imageSelector.ResourceNames)
            foreach (ISvgResource resourceName in entries)
            {
                // we store benchmarks objects directly in the listbox
                int index = lb.Items.Add(resourceName);
                if (resourceName.Equals(value))
                {
                    lb.SelectedIndex = index;
                }
            }

            // show this model stuff
            _editorService.DropDownControl(lb);
            if (lb.SelectedItem == null) // no selection, return the passed-in value as is
                return value;

            return lb.SelectedItem;
        }

        private static string GetStringForValue(object value)
        {
            if (value == null) return "null";
            return value.ToString();
        }

        //private void OnListBoxSelectedValueChanged(object sender, EventArgs e)
        //{
        //    // close the drop down as soon as something is clicked
        //    _editorService.CloseDropDown();
        //}
    }
}
