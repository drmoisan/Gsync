using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;


namespace Gsync.Utilities.Dialogs
{
    public enum BoxIcon
    {
        None = 0,
        Critical = 1,
        Warning = 2,
        Question = 4
    }

    public static class MyBox
    {
        public static DialogResult ShowDialog(string Message, string Title, BoxIcon icon, IList<DelegateButton> delegateButtons)
        {
            using (MyBoxViewer _viewer = new MyBoxViewer())
            {
                _viewer.Show();
                int columnWidth = _viewer.L2Bottom.GetColumnWidths()[1];
                _viewer.RemoveStandardButtons();
                _viewer.Text = Title;
                _viewer.TextMessage.Text = Message;
                

                Size tmp = _viewer.MinimumSize;

                foreach (var delegateButton in delegateButtons)
                {
                    AppendButtonInColumn(_viewer.L2Bottom, delegateButton, columnWidth);
                    tmp.Width += columnWidth;
                }

                _viewer.MinimumSize = tmp;
                _viewer.SetDialogIcon(icon);
                _viewer.Hide();
                DialogResult result = _viewer.ShowDialog();
                return result;
            }
        }

        public static DialogResult ShowDialog(MyBoxViewer viewer, string Message, string Title, BoxIcon icon, IList<ActionButton> actionButtons)
        {
            viewer.Show();
            ReplaceButtons(viewer, actionButtons);
            viewer.Text = Title;
            viewer.TextMessage.Text = Message;
            viewer.SetDialogIcon(icon);
            viewer.Hide();
            viewer.TopMost = true;
            DialogResult result = viewer.ShowDialog();
            return result;
        }

        public static T ShowDialog<T>(MyBoxViewer viewer, string Message, string Title, BoxIcon icon, FunctionButtonGroup<T> group)
        {
            viewer.Show();
            ReplaceButtons(viewer, group.FunctionButtons);
            viewer.Text = Title;
            viewer.TextMessage.Text = Message;
            viewer.SetDialogIcon(icon);
            viewer.Hide();
            viewer.TopMost = true;
            DialogResult result = viewer.ShowDialog();
            return group.Result;
        }

        public static DialogResult ShowDialog(MyBoxViewer viewer, string Message, string Title, MessageBoxIcon icon, IList<ActionButton> actionButtons)
        {
            viewer.Show();
            ReplaceButtons(viewer, actionButtons);
            viewer.Text = Title;
            viewer.TextMessage.Text = Message;
            viewer.SetDialogIcon(icon);
            viewer.Hide();
            viewer.TopMost = true;
            DialogResult result = viewer.ShowDialog();
            return result;
        }

        public static DialogResult ShowDialog(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon) 
        {
            var actionButtons = GetStandardButtons(buttons);
            using MyBoxViewer viewer = new();
            return ShowDialog(viewer, message, title, icon, actionButtons);
        }

        public static DialogResult ShowDialog(string message, string title, BoxIcon icon, Dictionary<string, Action> actions)
        {
            using MyBoxViewer viewer = new();
            var actionButtons = actions.ToActionButtons(viewer);
            return ShowDialog(viewer, message, title, icon, actionButtons);
        }

        public static T ShowDialog<T>(string message, string title, BoxIcon icon, Dictionary<string, Func<Task<T>>> functions)
        {
            using MyBoxViewer viewer = new();            
            var group = functions.ToFunctionButtonsAsync(viewer);
            return ShowDialog(viewer, message, title, icon, group);
        }

        internal static void ReplaceButtons(MyBoxViewer viewer, IList<ActionButton> actionButtons)
        {
            int columnWidth = viewer.L2Bottom.GetColumnWidths()[1];
            viewer.RemoveStandardButtons();

            Size minSize = viewer.MinimumSize;

            foreach (var actionButton in actionButtons)
            {
                AppendButtonInColumn(viewer.L2Bottom, actionButton, columnWidth);
                minSize.Width += columnWidth;
            }
            viewer.MinimumSize = minSize;
        }

        internal static void ReplaceButtons<T>(MyBoxViewer viewer, IList<FunctionButton<T>> functionButtons)
        {
            int columnWidth = viewer.L2Bottom.GetColumnWidths()[1];
            viewer.RemoveStandardButtons();

            Size minSize = viewer.MinimumSize;

            foreach (var functionButton in functionButtons)
            {
                AppendButtonInColumn(viewer.L2Bottom, functionButton, columnWidth);
                minSize.Width += columnWidth;
            }
            viewer.MinimumSize = minSize;
        }

        internal static IList<ActionButton> ToActionButtons(this Dictionary<string, Action> actions, MyBoxViewer _viewer)
        {
            IList<ActionButton> actionButtons = new List<ActionButton>();
            int i = 0;
            foreach (var actionPair in actions)
            {
                ActionButton actionButton;

                if (actionPair.Key.Contains("Cancel"))
                {
                    actionButton = new ActionButton(
                        $"button{i}", actionPair.Key, DialogResult.Cancel, actionPair.Value);
                }
                else
                {
                    actionButton = new ActionButton(
                        $"button{i}", actionPair.Key, DialogResult.OK, actionPair.Value);
                }

                actionButtons.Add(actionButton);
                i++;
            }
            return actionButtons;
        }

        public class FunctionButtonGroup<T>
        {
            public List<FunctionButton<T>> FunctionButtons { get; set; } = [];
            public T Result { get; set; }
        }

        internal static FunctionButtonGroup<T> ToFunctionButtonsAsync<T>(this Dictionary<string, Func<Task<T>>> functions, MyBoxViewer _viewer)
        {
            IList<FunctionButton<T>> functionButtons = new List<FunctionButton<T>>();
            int i = 0;
            var group = new FunctionButtonGroup<T>();
            foreach (var functionPair in functions)
            {                
                Func<Task<T>> function = async () =>
                {
                    var result = await functionPair.Value();
                    group.Result = result;
                    return result;
                };

                var functionButton = new FunctionButton<T>(
                    $"button{i}", functionPair.Key, DialogResult.OK, function);

                group.FunctionButtons.Add(functionButton);
                i++;
            }
            return group;
        }

        internal static void AppendButtonInColumn(TableLayoutPanel tlp, DelegateButton dlb, Single width)
        {
            tlp.ColumnCount++;
            tlp.ColumnStyles.Insert(tlp.ColumnCount-2, 
                                    new System.Windows.Forms.ColumnStyle(
                                        System.Windows.Forms.SizeType.Absolute,
                                        width));
            tlp.Controls.Add(dlb.Button, tlp.ColumnCount -2,0);
        }

        internal static void AppendButtonInColumn(TableLayoutPanel tlp, ActionButton actionButton, Single width)
        {
            tlp.ColumnCount++;
            tlp.ColumnStyles.Insert(tlp.ColumnCount - 2,
                                    new System.Windows.Forms.ColumnStyle(
                                        System.Windows.Forms.SizeType.Absolute,
                                        width));
            tlp.Controls.Add(actionButton.Button, tlp.ColumnCount - 2, 0);
        }

        internal static void AppendButtonInColumn<T>(TableLayoutPanel tlp, FunctionButton<T> functionButton, Single width)
        {
            tlp.ColumnCount++;
            tlp.ColumnStyles.Insert(tlp.ColumnCount - 2,
                                    new System.Windows.Forms.ColumnStyle(
                                        System.Windows.Forms.SizeType.Absolute,
                                        width));
            tlp.Controls.Add(functionButton.Button, tlp.ColumnCount - 2, 0);
        }

        private static void SetDialogIcon(this MyBoxViewer viewer, MessageBoxIcon icon)
        {
            switch (icon)
            {
                case MessageBoxIcon.None:
                    viewer.SvgIcon.Visible = false;
                    break;
                case MessageBoxIcon.Error:
                    viewer.SvgIcon.Image = SystemIcons.Error.ToBitmap();
                    break;
                case MessageBoxIcon.Warning:
                    // Same as MessageBoxIcon.Exclamation
                    viewer.SvgIcon.Image = SystemIcons.Warning.ToBitmap();
                    break;
                case MessageBoxIcon.Question:
                    viewer.SvgIcon.Image = SystemIcons.Question.ToBitmap();
                    break;
                case MessageBoxIcon.Information:
                    // Same as MessageBoxIcon.Asterisk
                    viewer.SvgIcon.Image = SystemIcons.Information.ToBitmap();
                    break;
                default:
                    break;
            }
        }


        private static void SetDialogIcon(this MyBoxViewer viewer, BoxIcon icon)
        {
            switch (icon)
            {
                case BoxIcon.None:
                    viewer.SvgIcon.Visible = false;
                    break;
                case BoxIcon.Critical:
                    viewer.SvgIcon.Image = SystemIcons.Error.ToBitmap();
                    break;
                case BoxIcon.Warning:
                    viewer.SvgIcon.Image = SystemIcons.Warning.ToBitmap();
                    break;
                case BoxIcon.Question:
                    viewer.SvgIcon.Image = SystemIcons.Question.ToBitmap();
                    break;
                default: break;
            }
        }

        internal static IList<ActionButton> GetStandardButtons(MessageBoxButtons buttons)
        {
            List<ActionButton> actionButtons = [];
            switch (buttons) 
            { 
                case MessageBoxButtons.OK:
                    actionButtons = [new ActionButton("ButtonOk", "Ok", DialogResult.OK, () => { })];
                    break;
                case MessageBoxButtons.OKCancel:
                    actionButtons = 
                        [
                            new ActionButton("ButtonOk", "Ok", DialogResult.OK, () => { }),
                            new ActionButton("ButtonCancel", "Cancel", DialogResult.Cancel, () => { })
                        ];
                    break;
                case MessageBoxButtons.AbortRetryIgnore:
                    actionButtons =
                        [
                            new ActionButton("ButtonAbort", "Abort", DialogResult.Abort, () => { }),
                            new ActionButton("ButtonRetry", "Retry", DialogResult.Retry, () => { }),
                            new ActionButton("ButtonIgnore", "Ignore", DialogResult.Ignore, () => { })
                        ];
                    break;
                case MessageBoxButtons.YesNoCancel:
                    actionButtons =
                        [
                            new ActionButton("ButtonYes", "Yes", DialogResult.Yes, () => { }),
                            new ActionButton("ButtonNo", "No", DialogResult.No, () => { }),
                            new ActionButton("ButtonCancel", "Cancel", DialogResult.Cancel, () => { })
                        ];
                    break;
                case MessageBoxButtons.YesNo:
                    actionButtons =
                        [
                            new ActionButton("ButtonYes", "Yes", DialogResult.Yes, () => { }),
                            new ActionButton("ButtonNo", "No", DialogResult.No, () => { }),
                        ];
                    break;
                case MessageBoxButtons.RetryCancel:
                    actionButtons =
                        [
                            new ActionButton("ButtonRetry", "Retry", DialogResult.Yes, () => { }),
                            new ActionButton("ButtonCancel", "Cancel", DialogResult.Cancel, () => { })
                        ];
                    break;
                default:
                    break;
            }
            return actionButtons;
        }

    }
}