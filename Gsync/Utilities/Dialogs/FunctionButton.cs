using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.Remoting.Contexts;
using System.Runtime.CompilerServices;
using Gsync.Utilities.Extensions;

namespace Gsync.Utilities.Dialogs
{
    public class FunctionButton<T>
    {
        private log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public FunctionButton() { }

        public FunctionButton(
            Button button,
            DialogResult dialogResult,
            Func<T> function)
        {
            _button = button;
            _button.DialogResult = dialogResult;
            _button.Click += new System.EventHandler(Button_Click);
            _function = function;
            _buttonClicked = function;
        }

        public FunctionButton(
            string name,
            string buttonText,
            Func<T> function)
        {
            _name = name;
            _function = function;            
            Button = MakeButton(buttonText);
            ButtonClicked = function;
        }

        public FunctionButton(
            string name,
            string buttonText,
            Func<T> function,
            Button template)
        {
            _template = template.Clone();
            _name = name;
            _function = function;
            ButtonClicked = function;
            Button = MakeButton(buttonText);
        }

        public FunctionButton(
            string name,
            string buttonText,
            DialogResult dialogResult,
            Func<T> function)
        {
            _name = name;
            _function = function;
            Button = MakeButton(buttonText, dialogResult);
            ButtonClicked = function;
        }

        public FunctionButton(
            string name,
            string buttonText,
            DialogResult dialogResult,
            Func<T> function,
            Button template)
        {
            _template = template.Clone();
            _name = name;
            _function = function;
            Button = MakeButton(buttonText, dialogResult);
            ButtonClicked = function;
        }

        public FunctionButton(
            string name,
            Image buttonImage,
            string buttonText,
            DialogResult dialogResult,
            Func<T> function)
        {
            _name = name;
            _function = function;
            Button = MakeButton(buttonText, buttonImage, dialogResult);
            ButtonClicked = function;
        }

        public FunctionButton(
            string name,
            Image buttonImage,
            string buttonText,
            DialogResult dialogResult,
            Func<T> function,
            Button template)
        {
            _template = template.Clone();
            _name = name;
            _function = function;
            Button = MakeButton(buttonText, buttonImage, dialogResult);
            ButtonClicked = function;
        }

        public FunctionButton(
            Button button,
            DialogResult dialogResult,
            Func<Task<T>> function)
        {
            _button = button;
            _button.DialogResult = dialogResult;
            _button.Click += new System.EventHandler(Button_Click);
            //_function = function;
            ButtonClickedAsync = function;
        }

        public FunctionButton(
            string name,
            string buttonText,
            Func<Task<T>> function)
        {
            _name = name;
            //_function = function;
            Button = MakeButton(buttonText);
            ButtonClickedAsync = function;
        }

        public FunctionButton(
            string name,
            string buttonText,
            Func<Task<T>> function,
            Button template)
        {
            _template = template.Clone();
            _name = name;
            //_function = function;
            ButtonClickedAsync = function;
            Button = MakeButton(buttonText);
        }

        public FunctionButton(
            string name,
            string buttonText,
            DialogResult dialogResult,
            Func<Task<T>> function)
        {
            _name = name;
            //_function = function;
            Button = MakeButton(buttonText, dialogResult);
            ButtonClickedAsync = function;
        }

        public FunctionButton(
            string name,
            string buttonText,
            DialogResult dialogResult,
            Func<Task<T>> function,
            Button template)
        {
            _template = template.Clone();
            _name = name;
            //_function = function;
            Button = MakeButton(buttonText, dialogResult);
            ButtonClickedAsync = function;
        }

        public FunctionButton(
            string name,
            Image buttonImage,
            string buttonText,
            DialogResult dialogResult,
            Func<Task<T>> function)
        {
            _name = name;
            Button = MakeButton(buttonText, buttonImage, dialogResult);
            ButtonClickedAsync = function;
        }

        public FunctionButton(
            string name,
            Image buttonImage,
            string buttonText,
            DialogResult dialogResult,
            Func<Task<T>> function,
            Button template)
        {
            _template = template.Clone();
            _name = name;            
            Button = MakeButton(buttonText, buttonImage, dialogResult);
            ButtonClickedAsync = function;
        }

        private string _name;
        private Button _button;
        private Button _template = new DelegateButtonTemplate().Button1; //.Clone();
        private Func<T> _function;


        public static FunctionButton<T> FromButton(
            Button button,
            DialogResult dialogResult,
            Func<T> function)
        {
            var db = new FunctionButton<T>();
            db.Button = button;
            db.Button.DialogResult = dialogResult;
            db.Delegate = function;

            return db;
        }

        public string Name { get => _name; set => _name = value; }

        public Button Button
        {
            get => _button;
            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_button != null)
                {
                    if (ButtonClicked is not null) { _button.Click -= Button_Click; }
                    if (ButtonClickedAsync is not null) { _button.Click -= Button_ClickAsync; }
                }                    
                _button = value;
                {
                    if (ButtonClicked is not null) { _button.Click += Button_Click; }
                    if (ButtonClickedAsync is not null) { _button.Click += Button_ClickAsync; }
                }
                
            }
        }

        public Func<T> Delegate { get => _function; set => _function = value; }

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

        private Func<T> _buttonClicked;
        public Func<T> ButtonClicked 
        { 
            get => _buttonClicked;
            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_buttonClicked is not null && _button is not null) { _button.Click -= Button_Click; }
                _buttonClicked = value;
                if (_buttonClicked is not null && _button is not null) { _button.Click += Button_Click; }                    
            }
        }

        private Func<Task<T>> _buttonClickedAsync;
        public Func<Task<T>> ButtonClickedAsync 
        { 
            get => _buttonClickedAsync;
            set
            {
                if (_buttonClickedAsync is not null && _button is not null) { _button.Click -= Button_ClickAsync; }
                _buttonClickedAsync = value;
                if (_buttonClickedAsync is not null && _button is not null) { _button.Click += Button_ClickAsync; }
            }
        }
        
        public T Value { get; internal set; }

        internal void Button_Click(object sender, EventArgs e)
        {
            try
            {
                Value = ButtonClicked.Invoke();
            }
            catch (Exception ex)
            {
                logger.Error($"Error in Button_Click: {ex.Message}", ex);
                throw;
            }
            
        }
        internal async void Button_ClickAsync(object sender, EventArgs e)
        {
            try
            {
                Value = await ButtonClickedAsync();
            }
            catch (Exception ex)
            {
                logger.Error($"Error in Button_ClickAsync: {ex.Message}", ex);
                throw;
            }
            
        }
    
    }
}
