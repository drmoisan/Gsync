using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Gsync.Utilities.Dialogs
{
        
    public enum YesNoToAllResponse
    {
        Empty = 0,
        Yes = 1,
        No = 2,
        YesToAll = 4,
        NoToAll = 8
    }


    public static class YesNoToAll
    {
        internal delegate void YesNoToAllDelegate();
        internal static void RespondYes() { Response = YesNoToAllResponse.Yes; }
        internal static void RespondYesToAll() { Response = YesNoToAllResponse.YesToAll; }
        internal static void RespondNo() { Response = YesNoToAllResponse.No; }
        internal static void RespondNoToAll() { Response = YesNoToAllResponse.NoToAll; }
        internal static void RespondCancel() { Response = YesNoToAllResponse.Empty; }

        private static YesNoToAllResponse _response = YesNoToAllResponse.Empty;
        public static YesNoToAllResponse Response { get => _response; set => _response = value; }

        public static YesNoToAllResponse ShowDialog(string message)
        {
            _response = YesNoToAllResponse.Empty;

            List<DelegateButton> delegateButtons = new List<DelegateButton>()
            { 
                new DelegateButton("Yes",Properties.Resources.Run1,"Yes",
                    DialogResult.OK, new YesNoToAllDelegate(YesNoToAll.RespondYes)),
                new DelegateButton("YesToAll",Properties.Resources.RepeatLastRun1,"YesToAll",
                    DialogResult.OK, new YesNoToAllDelegate(YesNoToAll.RespondYesToAll)),
                new DelegateButton("No",Properties.Resources.Cancel1,"No",
                    DialogResult.OK, new YesNoToAllDelegate(YesNoToAll.RespondNo)),
                new DelegateButton("NoToAll",Properties.Resources.RepeatUntilFailure1,"NoToAll",
                    DialogResult.OK, new YesNoToAllDelegate(YesNoToAll.RespondNoToAll)),
                new DelegateButton("Cancel","NoToAll",
                    DialogResult.Cancel, new YesNoToAllDelegate(YesNoToAll.RespondCancel))
            };
            MyBox.ShowDialog(message, "Dialog",BoxIcon.Question, delegateButtons);

            return Response;
        }
    }
}
