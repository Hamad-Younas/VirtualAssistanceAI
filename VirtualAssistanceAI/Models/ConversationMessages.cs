using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using VirtualAssistanceAI.Helpers;

namespace VirtualAssistanceAI.Models
{
    public class ConversationMessages : BaseHandler
    {
        public string MessageStatus { get; set; }

        public MessageType MessageType { get; set; }

        public ScriptName ScriptName { get; set; }

        public List<string> MessageResponse { get; set; }
        public string TimeStamp { get; set; }

        private string message;
        public string Message
        {
            get { return message; }
            set { message = value; NotifyPropertyChanged("Message"); }
        }
    }

    public enum MessageType
    {
        None,
        Issue,
        Confirmation,
        InputConfirmation,
        PasswordManager,
        Normal,
        RunpowerShell,
        PowerShellOptimizieChoice,
        PowerShellYourComputer,
        CPUIntensiveApps,
        Wait
    }

    public enum ScriptName
    {
        None,
        GetPrinterList,
        CheckPrinterExists,
        ClearPrintQueue,
        PrintTestPage,
        RestartPrinter,
        ResetPassword,
        RestartComputer,
        AskToUserAdjustBrightness,
        AskToUserWifiShare,
        UpdateSepcificApp,
        AskStillNoSound,
        PasswordManager,
        DidHearSound,
        UrlValidation,
        ChangeWifiPassword,
        CPUIntensiveApps,
    }
}
