using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gsync.OutlookInterop.Extensions;
using Gsync.Utilities.Extensions;


namespace Gsync.Utilities.Dialogs
{
    public class DelegateButton
    {
        public DelegateButton() { }
        
        public DelegateButton(Button button,
                              DialogResult dialogResult,
                              Delegate @delegate)
        {
            _button = button;
            _button.DialogResult = dialogResult;
            _delegate = @delegate;
        }

        public DelegateButton(string name,
                              string buttonText,
                              Delegate @delegate)
        {
            _name = name;
            _delegate = @delegate;
            Button = MakeButton(buttonText);
        }

        public DelegateButton(string name,
                              string buttonText,
                              Delegate @delegate,
                              Button template)
        {
            _template = template.Clone();
            _name = name;
            _delegate = @delegate;
            Button = MakeButton(buttonText);
        }

        public DelegateButton(string name,
                              string buttonText,
                              DialogResult dialogResult,
                              Delegate @delegate)
        {
            _name = name;
            _delegate = @delegate;
            Button = MakeButton(buttonText, dialogResult);
        }

        public DelegateButton(string name,
                              string buttonText,
                              DialogResult dialogResult,
                              Delegate @delegate,
                              Button template)
        { 
            _template = template.Clone();
            _name = name;
            _delegate = @delegate;
            Button = MakeButton(buttonText, dialogResult);
        }

        public DelegateButton(string name, 
                              Image buttonImage, 
                              string buttonText, 
                              DialogResult dialogResult, 
                              Delegate @delegate)
        {
            _name = name;
            _delegate = @delegate;
            Button = MakeButton(buttonText, buttonImage, dialogResult);
        }

        public DelegateButton(string name,
                              Image buttonImage,
                              string buttonText,
                              DialogResult dialogResult,
                              Delegate @delegate,
                              Button template)
        {
            _template = template.Clone();
            _name = name;
            _delegate = @delegate;
            Button = MakeButton(buttonText, buttonImage, dialogResult);
        }

        private string _name;
        private Button _button;
        private Button _template = new DelegateButtonTemplate().Button1.Clone();
        private Delegate _delegate;

        
        public static DelegateButton FromButton(Button button,
                                                DialogResult dialogResult,
                                                Delegate @delegate) 
        { 
            var db = new DelegateButton();
            db.Button = button;
            db.Button.DialogResult = dialogResult;
            db.Delegate = @delegate;

            return db; 
        }
        
        public string Name { get => _name; set => _name = value; }

        public Button Button
        {
            get => _button;
            set
            {
                if (_button != null)
                    _button.Click -= new System.EventHandler(Button_Click);
                _button = value;
                _button.Click += new System.EventHandler(Button_Click);
            }
        }
        
        public Delegate Delegate { get => _delegate; set => _delegate = value; }

        public Button ButtonTemplate { get => _template; set => _template = value.Clone(); }

        public Button MakeButton(string text)
        {
            Button b = _template.Clone();
            b.Visible = true;
            b.Enabled = true;
            b.Text = text;
            return b;
        }

        public Button MakeButton(string Text, Image Image)
        {
            Button button = MakeButton(Text);
            AddImageToButton(Image, ref button);
            return button;
        }

        public Button MakeButton(string Text, Image Image, DialogResult dialogResult)
        {
            Button button = MakeButton(Text, Image);
            button.DialogResult = dialogResult;
            return button;
        }

        public Button MakeButton(string Text, DialogResult dialogResult)
        {
            Button button = MakeButton(Text);
            button.DialogResult = dialogResult;
            return button;
        }

        private static void AddImageToButton(Image Image, ref Button button)
        {
            if (button.Image != null)
                button.Image.Dispose();
            button.Image = Image;
            button.TextImageRelation = TextImageRelation.ImageBeforeText;
        }

        internal void Button_Click(object sender, EventArgs e)
        {
            _delegate.DynamicInvoke();
        }
        
    }
}
