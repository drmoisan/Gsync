using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gsync.Utilities.Dialogs
{
    public partial class MyBoxViewer : Form
    {
        
        private bool ableToRemoveStandard = true;
        private readonly Dictionary<string, Delegate> _map;
                
        public MyBoxViewer()
        {
            InitializeComponent();
        }

        public MyBoxViewer(string title, string message, Dictionary<string, Delegate> map)
        {
            InitializeComponent();
            this.Text = title;
            this.TextMessage.Text = message;
            List<string> keys = map.Keys.ToList();
            Button1.Text = keys[0];
            Button2.Text = keys[1];
            this._map = map;
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            string key = _map.Keys.ToList()[0];
            var result = _map[key].DynamicInvoke();
            this.DialogResult = (DialogResult)result;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            string key = _map.Keys.ToList()[1];
            var result = _map[key].DynamicInvoke();
            this.DialogResult = (DialogResult)result;
        }

        public void RemoveStandardButtons()
        {
            if (ableToRemoveStandard)
            {
                RemoveStandardButtonControls();
                Size minSize = CalcMinSize();
                RemoveStandardButtonColumns();
                this.MinimumSize = minSize;
                ableToRemoveStandard = false;
            }
        }

        private void RemoveStandardButtonColumns()
        {
            this.L2Bottom.ColumnStyles.RemoveAt(2);
            this.L2Bottom.ColumnStyles.RemoveAt(1);
            this.L2Bottom.ColumnCount -= 2;
        }

        private void RemoveStandardButtonControls()
        {
            this.Button1.Click -= new System.EventHandler(this.Button1_Click);
            this.Button2.Click -= new System.EventHandler(this.Button2_Click);
            this.L2Bottom.Controls.Remove(this.Button1);
            this.L2Bottom.Controls.Remove(this.Button2);
            this.Button1.Dispose();
            this.Button2.Dispose();
        }

        internal Size CalcMinSize()
        {
            Size minSize = this.MinimumSize;
            float columnWidths = this.L2Bottom.ColumnStyles[2].Width;
            columnWidths += this.L2Bottom.ColumnStyles[1].Width;
            int columnWidthsRounded = (int)Math.Round(columnWidths, 0);
            if (minSize.Width > columnWidthsRounded)
            {
                minSize.Width -= columnWidthsRounded;
            }
            else
            {
                minSize.Width = 0;
            }

            return minSize;
        }

        private void GrowTextbox() 
        { 
            //var textWidth = TextRenderer.MeasureText(TextMessage.Text, TextMessage.Font).Width;
            var textHeight = TextRenderer.MeasureText(TextMessage.Text, TextMessage.Font, TextMessage.Size, TextFormatFlags.WordBreak).Height;
            if (textHeight > TextMessage.Height)
            {
                var textUsage = (double)textHeight * TextMessage.Width;
                var textArea = (double)TextMessage.Height * TextMessage.Width;
                var ratio = textUsage / textArea;
                var newTextSize = new Size((int)(ratio * TextMessage.Width), (int)(ratio * TextMessage.Height));
                var sizeDiff = newTextSize - TextMessage.Size ;
                this.Size += sizeDiff;
                this.MinimumSize += sizeDiff;
                this.TextMessage.Size = newTextSize;
            }
        }
        
        private void TextMessage_TextChanged(object sender, EventArgs e)
        {
            GrowTextbox();
        }
    }
}
