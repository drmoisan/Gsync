using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Gsync.Utilities.Extensions;


namespace Gsync.Utilities.Dialogs
{
    public class ActionButton
    {
        public ActionButton() { }

        public ActionButton(Button button,
                            DialogResult dialogResult,
                            Action action)
        {
            _button = button;
            _button.DialogResult = dialogResult;
            _button.Click += new System.EventHandler(Button_Click);
            _action = action;
        }

        public ActionButton(string name,
                            string buttonText,
                            Action action)
        {
            _name = name;
            _action = action;
            Button = MakeButton(buttonText);
        }

        public ActionButton(string name,
                            string buttonText,
                            Action action,
                            Button template)
        {
            _template = template.Clone();
            _name = name;
            _action = action;
            Button = MakeButton(buttonText);
        }

        public ActionButton(string name,
                            string buttonText,
                            DialogResult dialogResult,
                            Action action)
        {
            _name = name;
            _action = action;
            Button = MakeButton(buttonText, dialogResult);
        }

        public ActionButton(string name,
                            string buttonText,
                            DialogResult dialogResult,
                            Action action,
                            Button template)
        {
            _template = template.Clone();
            _name = name;
            _action = action;
            Button = MakeButton(buttonText, dialogResult);
        }

        public ActionButton(string name,
                            Image buttonImage,
                            string buttonText,
                            DialogResult dialogResult,
                            Action action)
        {
            _name = name;
            _action = action;
            Button = MakeButton(buttonText, buttonImage, dialogResult);
        }

        public ActionButton(string name,
                            Image buttonImage,
                            string buttonText,
                            DialogResult dialogResult,
                            Action action,
                            Button template)
        {
            _template = template.Clone();
            _name = name;
            _action = action;
            Button = MakeButton(buttonText, buttonImage, dialogResult);
        }

        private string _name;
        private Button _button;
        private Button _template = new DelegateButtonTemplate().Button1; //.Clone();
        private Action _action;


        public static ActionButton FromButton(Button button,
                                              DialogResult dialogResult,
                                              Action action)
        {
            var db = new ActionButton();
            db.Button = button;
            db.Button.DialogResult = dialogResult;
            db.Delegate = action;

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

        public Action Delegate { get => _action; set => _action = value; }

        public Button ButtonTemplate { get => _template; set => _template = value.Clone(); }

        public Button MakeButton(string text)
        {
            Button b = _template.Clone();
            b.Name = _name;
            b.Text = text;
            b.Visible = true;
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
            _action.DynamicInvoke();
        }

    }
}
