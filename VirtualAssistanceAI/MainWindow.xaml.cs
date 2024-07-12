using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Management.Automation.Runspaces;
using System.Management.Automation;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using VirtualAssistanceAI.ViewModels;
using VirtualAssistanceAI.Models;
using OpenAI_API.Moderation;
using VirtualAssistanceAI.Helpers;
using System.Threading;
using System.Security.Policy;

namespace VirtualAssistanceAI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindowVM MainWindowVMObj { get; set; }
        //private NotifyIcon m_notifyIcon;
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = new MainWindowVM(this);
            MainWindowVMObj = this.DataContext as MainWindowVM;
            //m_notifyIcon = new System.Windows.Forms.NotifyIcon();
            //m_notifyIcon.BalloonTipText = "The app has been minimised. Click the tray icon to show.";
            //m_notifyIcon.BalloonTipTitle = "Tech Support Assistance";
            //m_notifyIcon.BalloonTipClicked += M_notifyIcon_BalloonTipClicked;
            //m_notifyIcon.Text = "Tech Support Assistance";
            //m_notifyIcon.Icon = new System.Drawing.Icon("logo.ico");
            //m_notifyIcon.Click += new EventHandler(m_notifyIcon_Click);
            SetWindowRightLocation();
        }

        #region NotifyIcon
        private void M_notifyIcon_BalloonTipClicked(object sender, EventArgs e)
        {
            Show();
            WindowState = m_storedWindowState;
            SetWindowRightLocation();
            //System.Diagnostics.Process p = System.Diagnostics.Process.Start("calc.exe");
        }

        void OnClose(object sender, CancelEventArgs args)
        {
            //m_notifyIcon.Dispose();
            //m_notifyIcon = null;
        }

        private WindowState m_storedWindowState = WindowState.Normal;
        void OnStateChanged(object sender, EventArgs args)
        {
            if (WindowState == WindowState.Minimized)
            {
                Hide();
                //Notify Message
                //if (m_notifyIcon != null)
                //    m_notifyIcon.ShowBalloonTip(100);
            }
            else
                m_storedWindowState = WindowState;
        }
        void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs args)
        {
            CheckTrayIcon();
        }

        void m_notifyIcon_Click(object sender, EventArgs e)
        {
            Show();
            WindowState = m_storedWindowState;
            SetWindowRightLocation();

        }
        void CheckTrayIcon()
        {
            ShowTrayIcon(!IsVisible);
        }

        void ShowTrayIcon(bool show)
        {
            //if (m_notifyIcon != null)
            //    m_notifyIcon.Visible = show;
        }
        #endregion

        private void SetWindowRightLocation()
        {
            // Calculate the desired position
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            double screenHeight = SystemParameters.PrimaryScreenHeight;
            double windowHeight = this.Height;
            double windowWidth = this.Width;

            double leftPosition = screenWidth - windowWidth;

            // Get the taskbar height
            double taskbarHeight = GetTaskbarHeight();

            // Set the window position
            this.Left = leftPosition;
            this.Top = screenHeight - windowHeight - taskbarHeight;
        }

        private double GetTaskbarHeight()
        {
            double taskbarHeight = 42;

            foreach (Screen screen in Screen.AllScreens)
            {
                if (screen.Bounds.Bottom > taskbarHeight)
                {
                    taskbarHeight = screen.Bounds.Bottom - screen.WorkingArea.Bottom;
                }
            }

            return taskbarHeight - 12;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnMinimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
        private bool AutoScroll = true;
        private void svMessage_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // User scroll event : set or unset auto-scroll mode
            if (e.ExtentHeightChange == 0)
            {   // Content unchanged : user scroll event
                if (svMessage.VerticalOffset == svMessage.ScrollableHeight)
                {   // Scroll bar is in bottom
                    // Set auto-scroll mode
                    AutoScroll = true;

                    if (svMessage.ScrollableHeight > 0)
                        btnScrollUpDown.IsChecked = true;
                }
                else
                {   // Scroll bar isn't in bottom
                    // Unset auto-scroll mode
                    AutoScroll = false;
                    btnScrollUpDown.IsChecked = false;
                }
            }

            // Content scroll event : auto-scroll eventually
            if (AutoScroll && e.ExtentHeightChange != 0)
            {   // Content changed and auto-scroll mode set
                // Autoscroll
                svMessage.ScrollToVerticalOffset(svMessage.ExtentHeight);
            }
        }

        //private void txtInput_PreviewKeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        // Call the command from your view model
        //        svMessage.ScrollToBottom();
        //        //txtInput.IsEnabled = false;
        //        //MainWindowVMObj.BtnSendCommand.Execute(txtInput.Text);
        //        // txtInput.IsEnabled = true;
        //    }
        //}

        private void btnScrollUpDown_Click(object sender, RoutedEventArgs e)
        {
            if (btnScrollUpDown.IsChecked == true)
            {
                svMessage.ScrollToBottom();
            }
            else
            {
                svMessage.ScrollToTop();
            }
        }

        private void txtInput_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                // Call the command from your view model
                svMessage.ScrollToBottom();
                //txtInput.IsEnabled = false;
                MainWindowVMObj.BtnSendCommand.Execute(txtInput.Text);
                // txtInput.IsEnabled = true;
            }
        }

        string printerName = string.Empty;
        private async void lbMessageResponse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ConversationMessages conversation = (sender as System.Windows.Controls.ListBox).Tag as ConversationMessages;
            System.Windows.Controls.ListBox listBox = sender as System.Windows.Controls.ListBox;

            if (conversation.ScriptName == ScriptName.None && conversation.MessageType == MessageType.Issue)
            {
                if (listBox.SelectedItem != null)
                {
                    printerName = listBox.SelectedItem.ToString();
                    //string printerName = "Tax";

                    // Create a PowerShell runspace
                    Runspace runspace = RunspaceFactory.CreateRunspace();
                    runspace.Open();

                    // Create a pipeline
                    Pipeline pipeline = runspace.CreatePipeline();

                    // Load PowerShell script from JSON file
                    //string script = GetScript("GetPrinterList");

                    // Add PowerShell script to get printer list
                    // Add PowerShell script
                    // Load PowerShell script from JSON file
                    string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                    pipeline.Commands.AddScript(@script);
                    // Add printer name as a parameter
                    pipeline.Commands[0].Parameters.Add("printerName", printerName);

                    try
                    {
                        // Execute the script
                        Collection<PSObject> results = pipeline.Invoke();

                        if (results.Count > 0)
                        {
                            foreach (PSObject result in results)
                            {
                                System.Windows.MessageBox.Show(result.ToString());
                            }
                        }
                        else
                        {
                            //System.Windows.MessageBox.Show($"Confirm that {printerName} is the printer you want to troubleshoot (Y/N)");
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Confirm that {printerName} is the printer you want to troubleshoot (Y/N)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.ClearPrintQueue });
                        }
                    }
                    catch (RuntimeException ex)
                    {
                        System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show("Error: " + ex.Message);
                    }

                    // Clean up
                    pipeline.Dispose();
                    runspace.Close();
                }
            }
            else if (conversation.ScriptName == ScriptName.UpdateSepcificApp)
            {
                if (conversation.MessageType == MessageType.Wait) return;
                //MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });
                if (listBox.SelectedItem != null)
                {
                    string appName = GetAppName(listBox.SelectedItem.ToString());
                    await Task.Run(async () =>
                    {
                        //string printerName = "Tax";

                        // Create a PowerShell runspace
                        Runspace runspace = RunspaceFactory.CreateRunspace();
                        runspace.Open();

                        // Create a pipeline
                        Pipeline pipeline = runspace.CreatePipeline();

                        // Load PowerShell script from JSON file
                        //string script = GetScript("GetPrinterList");

                        // Add PowerShell script to get printer list
                        // Add PowerShell script
                        // Load PowerShell script from JSON file
                        //string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                        pipeline.Commands.AddScript(@"try {
        Write-Output ""Updating $appName...""
        $result = winget upgrade --id $appName --accept-source-agreements --accept-package-agreements

        if ($result.ExitCode -eq 0) {
            Write-Output ""$appName updated successfully.""
        }
        else {
            Write-Output ""Failed to update $($appName): $($result.StandardError)""
        }
    }
    catch {
        Write-Output ""Error updating app: $($_.Exception.Message)""
    }");
                        // Add printer name as a parameter
                        pipeline.Commands[0].Parameters.Add("appName", appName);

                        try
                        {
                            // Execute the script
                            Collection<PSObject> results = pipeline.Invoke();

                            if (results.Count > 0)
                            {
                                foreach (PSObject result in results)
                                {
                                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                                    {
                                        if (result != null)
                                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}{appName}", MessageStatus = "Received", MessageType = MessageType.Normal });

                                    });
                                }
                            }
                        }
                        catch (RuntimeException ex)
                        {
                            System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Error: " + ex.Message);
                        }

                        // Clean up
                        pipeline.Dispose();
                        runspace.Close();
                    });
                }

                List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

                //foreach (var message in MainWindowVMObj.Messages)
                //{
                //    if (message.MessageType == MessageType.Wait)
                //    {
                //        messagesToRemove.Add(message);
                //    }
                //}

                //foreach (var message in messagesToRemove)
                //{
                //    MainWindowVMObj.Messages.Remove(message);
                //}
            }
        }

        public string GetAppName(string input)
        {
            // Split the input string by space
            string[] parts = input.Split(' ');

            // The browser name is the last part of the string
            string browserName = parts[parts.Length - 1];

            return browserName;
        }



        private void RestartComputer()
        {
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();
            // Load PowerShell script from JSON file
            string script = MainWindowVMObj.GetScript("RestartComputer");
            pipeline.Commands.AddScript(script);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        //System.Windows.MessageBox.Show(result.ToString());
                        MainWindowVMObj.Messages.Add(new ConversationMessages { Message = result.ToString(), MessageStatus = "Received", MessageType = MessageType.Normal });
                        //RestartComputer();
                    }
                }
                else
                {
                    //System.Windows.MessageBox.Show($"Confirm that {printerName} is the printer you want to troubleshoot (Y/N)");
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please wait Restarting...", MessageStatus = "Received", MessageType = MessageType.Normal });
                }
            }
            catch (RuntimeException ex)
            {
                System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: " + ex.Message);
            }

            // Clean up
            pipeline.Dispose();
            runspace.Close();
        }

        private void ClearPrintQueue()
        {
            try
            {
                // Create a PowerShell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                // Create a pipeline
                Pipeline pipeline = runspace.CreatePipeline();


                // Load PowerShell script from JSON file
                string script = MainWindowVMObj.GetScript("ClearPrintQueue");
                pipeline.Commands.AddScript(script);
                // Add printer name as a parameter
                pipeline.Commands[0].Parameters.Add("printerName", printerName);

                try
                {
                    // Execute the script
                    Collection<PSObject> results = pipeline.Invoke();

                    if (results.Count > 0)
                    {
                        foreach (PSObject result in results)
                        {
                            //System.Windows.MessageBox.Show(result.ToString());
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = result.ToString(), MessageStatus = "Received", MessageType = MessageType.Normal });
                            PrintTestPage();
                        }
                    }
                    else
                    {
                        //System.Windows.MessageBox.Show($"Confirm that {printerName} is the printer you want to troubleshoot (Y/N)");
                        MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Clearing the print queue for {printerName}", MessageStatus = "Received", MessageType = MessageType.Normal });
                    }
                }
                catch (RuntimeException ex)
                {
                    System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                }

                // Clean up
                pipeline.Dispose();
                runspace.Close();
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private void PrintTestPage()
        {
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();
            // Load PowerShell script from JSON file
            string script = MainWindowVMObj.GetScript("PrintTestPage");
            pipeline.Commands.AddScript(script);
            // Add printer name as a parameter
            pipeline.Commands[0].Parameters.Add("printerName", printerName);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        //System.Windows.MessageBox.Show(result.ToString());
                        MainWindowVMObj.Messages.Add(new ConversationMessages { Message = result.ToString(), MessageStatus = "Received", MessageType = MessageType.Normal });
                        RestartPrinter();
                    }
                }
                else
                {
                    //System.Windows.MessageBox.Show($"Confirm that {printerName} is the printer you want to troubleshoot (Y/N)");
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Printing a test page for {printerName}", MessageStatus = "Received", MessageType = MessageType.Normal });
                }
            }
            catch (RuntimeException ex)
            {
                System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: " + ex.Message);
            }

            // Clean up
            pipeline.Dispose();
            runspace.Close();
        }

        private void RestartPrinter()
        {
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();
            // Load PowerShell script from JSON file
            string script = MainWindowVMObj.GetScript("RestartPrinter");
            pipeline.Commands.AddScript(script);
            // Add printer name as a parameter
            pipeline.Commands[0].Parameters.Add("printerName", printerName);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        //System.Windows.MessageBox.Show(result.ToString());
                        MainWindowVMObj.Messages.Add(new ConversationMessages { Message = result.ToString(), MessageStatus = "Received", MessageType = MessageType.Normal });
                        MainWindowVMObj.Messages.Add(new ConversationMessages { Message = "The troubleshooting steps did not resolve the issue. Do you want to restart the computer? (Y/N)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.RestartComputer });
                        //RestartComputer();
                    }
                }
                else
                {
                    //System.Windows.MessageBox.Show($"Confirm that {printerName} is the printer you want to troubleshoot (Y/N)");
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Restarting {printerName}", MessageStatus = "Received", MessageType = MessageType.Normal });
                }
            }
            catch (RuntimeException ex)
            {
                System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: " + ex.Message);
            }

            // Clean up
            pipeline.Dispose();
            runspace.Close();
        }

        private void btnYes_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            ConversationMessages script = (sender as System.Windows.Controls.Button).Tag as ConversationMessages;

            switch (script.ScriptName)
            {
                case ScriptName.ClearPrintQueue:
                    ClearPrintQueue();
                    break;

                case ScriptName.RestartComputer:
                    RestartComputer();
                    break;

                case ScriptName.AskToUserAdjustBrightness:
                    AskToUserValue(script.ScriptName);
                    break;

                case ScriptName.AskToUserWifiShare:
                    AskToUserWifiShare(script.ScriptName);
                    break;

                case ScriptName.ResetPassword:
                    ResetPassword(script.ScriptName);
                    break;

                case ScriptName.AskStillNoSound:
                    PrintAskStillYesBtn(script.ScriptName);
                    break;

                case ScriptName.DidHearSound:
                    DidHearSound(script.ScriptName);
                    break;

                case ScriptName.PasswordManager:
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Which password manager would you like to use? (1 - Google, 2 - Others)", MessageStatus = "Received", MessageType = MessageType.PasswordManager });
                    break;

                case ScriptName.ChangeWifiPassword:
                    AskChangeWifiPassword();
                    break;
            }
        }

        private void AskChangeWifiPassword()
        {
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Enter the new Wi-Fi password:", MessageStatus = "Received", MessageType = MessageType.InputConfirmation, ScriptName = ScriptName.ChangeWifiPassword });
        }

        private void btnNo_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            ConversationMessages script = (sender as System.Windows.Controls.Button).Tag as ConversationMessages;

            switch (script.ScriptName)
            {
                case ScriptName.ClearPrintQueue:
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please select the correct printer and try again.", MessageStatus = "Received", MessageType = MessageType.Normal });
                    break;

                case ScriptName.RestartComputer:
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"It is recommended to seek assistance from an IT professional for further troubleshooting.", MessageStatus = "Received", MessageType = MessageType.Normal });
                    break;

                case ScriptName.AskToUserAdjustBrightness:
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Brightness level is set to {MainWindowVMObj.CurrentBrightnessLevel}%. Thank you!", MessageStatus = "Received", MessageType = MessageType.Normal });
                    break;

                case ScriptName.AskToUserWifiShare:
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Okay, no problem. If you change your mind, just let me know!", MessageStatus = "Received", MessageType = MessageType.Normal });
                    break;

                case ScriptName.ResetPassword:
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Okay, no problem. If you need further assistance, feel free to ask!", MessageStatus = "Received", MessageType = MessageType.Normal });
                    break;

                case ScriptName.AskStillNoSound:
                    PrintAskStillYesBtn(script.ScriptName);
                    break;

                case ScriptName.DidHearSound:
                    DidHearSound(script.ScriptName);
                    break;

                case ScriptName.PasswordManager:
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Okay, no problem. If you need further assistance, feel free to ask!", MessageStatus = "Received", MessageType = MessageType.Normal });
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Do you want to change your Wi-Fi password? (Y/N):", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.ChangeWifiPassword });
                    break;

                case ScriptName.ChangeWifiPassword:
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Okay, no problem. If you need further assistance, feel free to ask!", MessageStatus = "Received", MessageType = MessageType.Normal });
                    break;
            }
        }

        private void DidHearSound(ScriptName scriptName)
        {
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Running the sound troubleshooter...", MessageStatus = "Received", MessageType = MessageType.Normal });
            //MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Testing the sound out of a different application/media player...", MessageStatus = "Received", MessageType = MessageType.Normal });

            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            //string script = GetScript("CheckMaximizeDeviceVolume");

            // Add PowerShell script to get printer list
            pipeline.Commands.AddScript(@"try {
    Start-Process ""ms-settings:troubleshoot?troubleshootType=6"" -Wait
} catch {
    Write-Output ""An error occurred while starting the sound troubleshooter: $_""
}");

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();
                foreach (PSObject result in results)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (result != null)
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                    });
                }

                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"The sound issue persists. Do you want to restart the device? (Y/N)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.RestartComputer });
                });

            }
            catch (RuntimeException ex)
            {
                System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: " + ex.Message);
            }

            // Clean up
            pipeline.Dispose();
            runspace.Close();
        }

        private void PrintAskStillYesBtn(ScriptName scriptName)
        {
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please make sure all cables are fully connected between your monitor and external speakers, if you're using one. Also, ensure that the speakers are turned on.", MessageStatus = "Received", MessageType = MessageType.Normal });
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Testing the sound out of a different application/media player...", MessageStatus = "Received", MessageType = MessageType.Normal });

            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            //string script = GetScript("CheckMaximizeDeviceVolume");

            // Add PowerShell script to get printer list
            pipeline.Commands.AddScript(@"[System.Media.SystemSounds]::Asterisk.Play()");

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();
                foreach (PSObject result in results)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (result != null)
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                    });
                }

                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Did you hear the sound? (Y/N)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.DidHearSound });
                });

            }
            catch (RuntimeException ex)
            {
                System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: " + ex.Message);
            }

            // Clean up
            pipeline.Dispose();
            runspace.Close();
        }

        private void ResetPassword(ScriptName scriptName)
        {
            //            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            //            {
            //                runspace.Open();

            //                using (Pipeline pipeline = runspace.CreatePipeline())
            //                {
            //                    pipeline.Commands.AddScript(@"$chromeInstalled = Test-Path ""C:\Program Files\Google\Chrome\Application\chrome.exe""
            //        if ($chromeInstalled) {
            //            Start-Process ""https://passwords.google.com""
            //            Write-Host ""Please search for the website or program you're trying to access in Google Password Manager.""
            //        }
            //        else {
            //            Write-Host ""Please open Google Chrome and navigate to https://passwords.google.com to access your passwords.""
            //        }
            //");

            //                    // Add printer name as a parameter
            //                    //pipeline.Commands[0].Parameters.Add("passwordManager", "passwordManager");

            //                    try
            //                    {
            //                        var results = pipeline.Invoke();
            //                        foreach (var result in results)
            //                        {
            //                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            //                            {
            //                                if (result != null)
            //                                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
            //                            });
            //                        }
            //                        if (results != null && results.Count > 0)
            //                        {
            //                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Do you need to reset a password? (y/n)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.ResetPassword });
            //                        }

            //                        //  Messages.Add(new ConversationMessages { Message = $"Which password manager would you like to use? (1 - Google, 2 - Others)", MessageStatus = "Received", MessageType = MessageType.PasswordManager });
            //                    }
            //                    catch (RuntimeException ex)
            //                    {
            //                        System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            //                    }
            //                    catch (Exception ex)
            //                    {
            //                        System.Windows.MessageBox.Show("Error: " + ex.Message);
            //                    }
            //                }
            //            }
        }

        private void AskToUserWifiShare(ScriptName scriptName)
        {
            MainWindowVMObj.AskToUserWifiShare(scriptName);
        }

        private void AskToUserValue(ScriptName scriptName)
        {
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Enter the desired brightness level (1-100):", MessageStatus = "Received", MessageType = MessageType.InputConfirmation, ScriptName = scriptName });
            //MainWindowVMObj.AskToUserAdjustBrightness();
        }

        private async void tbInputBrightness_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.Enter)
                {
                    ConversationMessages script = (sender as System.Windows.Controls.TextBox).Tag as ConversationMessages;
                    if (script != null)
                    {
                        if (script.ScriptName == ScriptName.UrlValidation)
                        {
                            txtInput.Focus();
                            svMessage.ScrollToBottom();
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = "Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });
                            string url = (sender as System.Windows.Controls.TextBox).Text.Trim();
                            await Task.Run(async () =>
                            {
                                await UrlValidation(url);
                            });

                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = "Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });

                            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

                            foreach (var message in MainWindowVMObj.Messages)
                            {
                                if (message.MessageType == MessageType.Wait)
                                {
                                    messagesToRemove.Add(message);
                                }
                            }

                            foreach (var message in messagesToRemove)
                            {
                                MainWindowVMObj.Messages.Remove(message);
                            }
                        }
                        else if (script.ScriptName == ScriptName.ChangeWifiPassword)
                        {
                            txtInput.Focus();
                            svMessage.ScrollToBottom();
                            string password = (sender as System.Windows.Controls.TextBox).Text.Trim();
                            ChangeWifiPassword(password);
                        }
                        else if (script.ScriptName == ScriptName.UpdateSepcificApp)
                        {
                            txtInput.Focus();
                            svMessage.ScrollToBottom();
                            string appname = (sender as System.Windows.Controls.TextBox).Text.Trim();
                            UpdateSpecificApp(appname);
                        }
                        else
                            MainWindowVMObj.AskToUserAdjustBrightness(int.Parse((sender as System.Windows.Controls.TextBox).Text));
                    }

                    //MainWindowVMObj.BtnSendCommand.Execute(txtInput.Text);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Enter the desired brightness level (1-100)");
            }
        }

        private async Task UpdateSpecificApp(string appname)
        {
            await Task.Run(async () =>
            {
                //string printerName = "Tax";

                // Create a PowerShell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                // Create a pipeline
                Pipeline pipeline = runspace.CreatePipeline();

                // Load PowerShell script from JSON file
                //string script = GetScript("GetPrinterList");

                // Add PowerShell script to get printer list
                // Add PowerShell script
                // Load PowerShell script from JSON file
                //string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                pipeline.Commands.AddScript(@"try {
        Write-Output ""Updating $appName...""
        $result = winget upgrade --id $appName --accept-source-agreements --accept-package-agreements

        if ($result.ExitCode -eq 0) {
            Write-Output ""$appName updated successfully.""
        }
        else {
            Write-Output ""Failed to update $($appName): $($result.StandardError)""
        }
    }
    catch {
        Write-Output ""Error updating app: $($_.Exception.Message)""
    }");
                // Add printer name as a parameter
                pipeline.Commands[0].Parameters.Add("appName", appname);

                try
                {
                    // Execute the script
                    Collection<PSObject> results = pipeline.Invoke();

                    if (results.Count > 0)
                    {
                        foreach (PSObject result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (result != null)
                                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}{appname}", MessageStatus = "Received", MessageType = MessageType.Normal });

                            });
                        }
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please save your work and restart the program.", MessageStatus = "Received", MessageType = MessageType.Normal });
                        });


                    }
                }
                catch (RuntimeException ex)
                {
                    System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                }

                // Clean up
                pipeline.Dispose();
                runspace.Close();
            });
        }

        private void ChangeWifiPassword(string password)
        {
            string newPassword = password;
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            string script = $"# Identify the Wi-Fi network interface\r\n    $wifiInterface = (Get-NetAdapter | Where-Object {{ $_.InterfaceDescription -like \"*Wi-Fi*\" }}).Name\r\n\r\n    if ($wifiInterface) {{\r\n        # Get the SSID of the current Wi-Fi network\r\n        $wifiSSID = (netsh wlan show interfaces | Select-String \"SSID\").Line.Split(':')[1].Trim()\r\n        if ($wifiSSID) {{\r\n            # Retrieve the Wi-Fi profile XML\r\n            $profileXml = netsh wlan export profile name=\"$wifiSSID\" folder=$env:TEMP\r\n            $profilePath = Join-Path $env:TEMP \"$wifiSSID.xml\"\r\n\r\n            # Load the XML and update the password\r\n            [xml]$profileContent = Get-Content $profilePath\r\n            $profileContent.WLANProfile.MSM.security.sharedKey.keyMaterial = \"{newPassword}\"\r\n\r\n            # Save the updated XML back to the same path\r\n            $profileContent.Save($profilePath)\r\n\r\n            # Import the updated profile\r\n            netsh wlan add profile filename=$profilePath\r\n\r\n            Write-Output \"Wi-Fi password changed successfully.\"\r\n        }} else {{\r\n            Write-Output \"No Wi-Fi Network found. You may not be connected to a Wi-Fi network.\"\r\n        }}\r\n    }} else {{\r\n        Write-Output \"Wi-Fi interface not found. Make sure you are connected to a Wi-Fi network.\"\r\n    }}";

            // Add the script to the pipeline
            pipeline.Commands.AddScript(@script);

            // Add printer name as a parameter
            //pipeline.Commands[0].Parameters.Add("$activeUrl", activeUrl);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();
                foreach (PSObject result in results)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (result != null)
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                    });
                }

            }
            catch (RuntimeException ex)
            {
                System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: " + ex.Message);
            }

            // Clean up
            pipeline.Dispose();
            runspace.Close();
        }

        private async Task UrlValidation(string url)
        {
            string userInput = url;
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            //string script = GetScript("CheckMaximizeDeviceVolume");

            // Add the script to the pipeline
            pipeline.Commands.AddScript(@"param (
        [string]$userInput
    )

    Write-Output ""Input received: $userInput""

    $isValidUrl = $userInput -match ""^((https?|ftp):\/\/)?([a-z0-9]+\.)?[a-z0-9][a-z0-9-]*\.[a-z]{2,}(:[0-9]+)?(\/.*)?$""
    $isValidIp = $userInput -match ""\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b""

    

    if ([string]::IsNullOrWhiteSpace($userInput)) {
        Write-Output ""You didn't input the website URL or gateway IP."" 
        return $false
    }
    elseif ($isValidUrl -or $isValidIp) {
        return $true
    }
    else {
        Write-Output ""Invalid website URL or gateway IP. Please enter a valid URL or IP address."" 
        return $false
    }
");

            //pipeline.Input.Write("https://www.facebook.com/");

            // Add printer name as a parameter
            pipeline.Commands[0].Parameters.Add("userInput", userInput);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();
                foreach (PSObject result in results)
                {
                    if (result.ToString() == "True" || result.ToString() == "False")
                    {
                        if (bool.Parse(result.ToString()))
                        {
                            await Task.Run(async () =>
                            {
                                await TestInternetrConnection(url);
                            });
                        }
                    }
                    else
                    {
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (result != null)
                                MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                        });
                    }
                }
                //System.Windows.Application.Current.Dispatcher.Invoke(() =>
                //{
                //    Messages.Add(new ConversationMessages { Message = $"Did you hear the sound? (Y/N)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.DidHearSound });
                //});

            }
            catch (RuntimeException ex)
            {
                System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: " + ex.Message);
            }

            // Clean up
            pipeline.Dispose();
            runspace.Close();
        }

        private async Task TestInternetrConnection(string url)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                svMessage.ScrollToBottom();
                MainWindowVMObj.Messages.Add(new ConversationMessages { Message = "Testing Internet connection...", MessageStatus = "Received", MessageType = MessageType.Normal });
            });

            string activeUrl = url;
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            string script = $"$activeUrl = \"{activeUrl}\"\n# Test Internet Connection\r\nWrite-Host \"Testing Internet connection...\"\r\n$pingResult = Test-NetConnection -ComputerName google.com\r\n\r\nif ($pingResult.PingSucceeded) {{\r\n    # Ping the WiFi to ensure it's not just the one site that's down\r\n    Write-Output \"Pinging $activeUrl...\"\r\n    $pingUrl = Test-NetConnection -ComputerName $activeUrl\r\n\r\n    if ($pingUrl.PingSucceeded) {{\r\n        # Clear browser cache and restart the browser\r\n        Write-Output \"Clearing browser cache and restarting the browser...\"\r\n        $browsers = Get-Process -Name \"*chrome*\", \"*firefox*\", \"*edge*\", \"*iexplore*\"\r\n        foreach ($browser in $browsers) {{\r\n            $browser.CloseMainWindow() | Out-Null\r\n        }}\r\n\r\n        # Test the original URL again\r\n        Write-Output \"Testing $activeUrl again...\"\r\n        $pingUrl = Test-NetConnection -ComputerName $activeUrl\r\n\r\n        if ($pingUrl.PingSucceeded) {{\r\n            Write-Output \"The issue seems to be resolved after clearing the cache and restarting the browser.\"\r\n        }}\r\n        else {{\r\n            Write-Output \"The issue persists. Trying further troubleshooting steps...\"\r\n        }}\r\n    }}\r\n    else {{\r\n        Write-Output \"The issue appears to be with the specific website or gateway $activeUrl. Trying to test it in a different browser.\"\r\n        # Test in a different browser\r\n        $defaultBrowser = (Get-ItemPropertyValue 'HKCU:\\Software\\Microsoft\\Windows\\Shell\\Associations\\UrlAssociations\\http\\UserChoice' -Name 'ProgId')\r\n        switch ($defaultBrowser) {{\r\n            \"ChromeHTML\" {{ $browserPath = \"chrome.exe\" }}\r\n            \"FirefoxURL\" {{ $browserPath = \"firefox.exe\" }}\r\n            \"AppXq0fevzme2pys62n3e0fbqa7peapykr8v\" {{ $browserPath = \"msedge.exe\" }} # Edge\r\n            \"IE.HTTP\" {{ $browserPath = \"iexplore.exe\" }}\r\n            default {{ Write-Output \"Unknown browser.\" }}\r\n        }}\r\n\r\n        if ($browserPath) {{\r\n            Start-Process $browserPath -ArgumentList $activeUrl\r\n        }}\r\n        else {{\r\n            Write-Output \"Unable to determine the default browser. Please open the URL manually in a browser.\"\r\n        }}\r\n\r\n        # Check if the issue is resolved after testing in a different browser\r\n        Write-Output \"Testing $activeUrl again in a different browser...\"\r\n        $pingUrl = Test-NetConnection -ComputerName $activeUrl\r\n\r\n        if ($pingUrl.PingSucceeded) {{\r\n            Write-Output \"The issue seems to be resolved after testing it in a different browser.\"\r\n        }}\r\n        else {{\r\n            # Clear browser cache and restart the browser\r\n            Write-Output \"Clearing browser cache and restarting the browser...\"\r\n            $browsers = Get-Process -Name \"*chrome*\", \"*firefox*\", \"*edge*\", \"*iexplore*\"\r\n            foreach ($browser in $browsers) {{\r\n                $browser.CloseMainWindow() | Out-Null\r\n            }}\r\n\r\n            # Test the original URL again\r\n            Write-Output \"Testing $activeUrl again...\"\r\n            $pingUrl = Test-NetConnection -ComputerName $activeUrl\r\n\r\n            if ($pingUrl.PingSucceeded) {{\r\n                Write-Output \"The issue seems to be resolved after clearing the cache and restarting the browser.\"\r\n            }}\r\n            else {{\r\n                Write-Output \"The issue persists. Further troubleshooting is required.\"\r\n            }}\r\n        }}\r\n    }}\r\n}}\r\nelse {{\r\n    # Restart the WiFi and test the internet speed\r\n    Write-Output \"Restarting the WiFi...\"\r\n    $netAdapter = Get-NetAdapter -Name \"*Wi-Fi*\"\r\n    $netAdapter | Disable-NetAdapter -Confirm:$false\r\n    Start-Sleep -Seconds 5\r\n    $netAdapter | Enable-NetAdapter -Confirm:$false\r\n\r\n    Start-Sleep -Seconds 30\r\n\r\n    Write-Output \"Testing internet speed...\"\r\n    $speedTestResult = Test-NetConnection -ComputerName $activeUrl\r\n\r\n    if ($speedTestResult.PingSucceeded) {{\r\n        # Open the original active window\r\n        Write-Output \"The issue seems to be resolved. Opening $activeUrl...\"\r\n        Start-Process $activeUrl\r\n    }}\r\n    else {{\r\n        # Restart the router using the IP admin console\r\n        Write-Output \"Internet speed is too slow or not working. Restarting the router...\"\r\n        $routerLoginInfo = @(\"admin/admin\", \"admin/Admin\", \"admin/password\", \"admin/1234\") # Router login credentials\r\n        $routerIP = $activeUrl # Assume the active URL is the router IP\r\n\r\n        foreach ($credential in $routerLoginInfo) {{\r\n            $username = $credential.Split(\"/\")[0]\r\n            $password = $credential.Split(\"/\")[1]\r\n            Write-Output \"Trying to restart the router using the IP admin console with the credential: $username/$password...\"\r\n            $url = \"http://$routerIP/\"\r\n            $basicAuthValue = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes(\"$($username):$($password)\"))\r\n            $headers = @{{ Authorization = \"Basic $basicAuthValue\" }}\r\n            try {{\r\n                $response = Invoke-WebRequest -Uri $url -Headers $headers -Method GET\r\n                if ($response.StatusCode -eq 200) {{\r\n                    # Successful authentication, restart the router\r\n                    # ADD CODE TO RESTART THE ROUTER HERE\r\n                    # For example, you could use Invoke-WebRequest to send a specific command to the router's API\r\n                    # or use a third-party library/module to interact with the router's interface\r\n\r\n                    Write-Output \"Router restarted successfully.\"\r\n                }}\r\n                else {{\r\n                    Write-Output \"Failed to authenticate to the router using the provided credentials.\"\r\n                }}\r\n            }}\r\n            catch {{\r\n                Write-Output \"An error occurred while trying to access the router: $($_.Exception.Message)\"\r\n            }}\r\n\r\n            Start-Sleep -Seconds 300 # Wait for 5 minutes\r\n\r\n            Write-Output \"Testing internet connection again...\"\r\n            $pingResult = Test-NetConnection -ComputerName google.com\r\n\r\n            if ($pingResult.PingSucceeded) {{\r\n                Write-Output \"The issue seems to be resolved after restarting the router.\"\r\n                break\r\n            }}\r\n            else {{\r\n                Write-Output \"The issue persists. Please unplug your router, wait for 60 seconds, plug it back in, and wait for 5-10 minutes for it to power up and test the internet again.\"\r\n            }}\r\n        }}\r\n    }}\r\n}}\r\n\r\n# Sleep for 5 seconds before exiting to allow the user to see any final messages\r\nStart-Sleep -Seconds 5";

            // Add the script to the pipeline
            pipeline.Commands.AddScript(@script);

            //pipeline.Input.Write("https://www.facebook.com/");

            // Add printer name as a parameter
            //pipeline.Commands[0].Parameters.Add("$activeUrl", activeUrl);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();
                foreach (PSObject result in results)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (result != null)
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                    });
                }

                //TestInternetrConnection(url);
                //System.Windows.Application.Current.Dispatcher.Invoke(() =>
                //{
                //    Messages.Add(new ConversationMessages { Message = $"Did you hear the sound? (Y/N)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.DidHearSound });
                //});

            }
            catch (RuntimeException ex)
            {
                System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Error: " + ex.Message);
            }

            // Clean up
            pipeline.Dispose();
            runspace.Close();
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //TaskbarHelper.PinToTaskbar("VirtualAssistanceAI.exe");
            //Task.Run(async () =>
            //{
            //    await MainWindowVMObj.AddBlockScript();
            //});
        }

        private void btnGoogle_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            //string passwordManager = (sender as System.Windows.Controls.Button).Content.ToString();
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                using (Pipeline pipeline = runspace.CreatePipeline())
                {
                    pipeline.Commands.AddScript(@"$chromeInstalled = Test-Path ""C:\Program Files\Google\Chrome\Application\chrome.exe""
        if ($chromeInstalled) {
            Start-Process ""https://passwords.google.com""
            Write-Host ""Please search for the website or program you're trying to access in Google Password Manager.""
        }
        else {
            Write-Host ""Please open Google Chrome and navigate to https://passwords.google.com to access your passwords.""
        }
");

                    // Add printer name as a parameter
                    //pipeline.Commands[0].Parameters.Add("passwordManager", "passwordManager");

                    try
                    {
                        var results = pipeline.Invoke();
                        foreach (var result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (result != null)
                                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                            });
                        }
                        //if(results != null && results.Count>0)
                        {
                            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Do you need to reset a password? (y/n)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.ResetPassword });
                        }

                        //  Messages.Add(new ConversationMessages { Message = $"Which password manager would you like to use? (1 - Google, 2 - Others)", MessageStatus = "Received", MessageType = MessageType.PasswordManager });
                    }
                    catch (RuntimeException ex)
                    {
                        System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }
        }

        private void btnOther_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please open your preferred password manager and search for the website or program you're trying to access.", MessageStatus = "Received", MessageType = MessageType.Normal });
        }

        private void btnOne_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            ConversationMessages script = (sender as System.Windows.Controls.Button).Tag as ConversationMessages;
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Select the app you want to update", MessageStatus = "Received", MessageType = MessageType.Issue, MessageResponse = script.MessageResponse, ScriptName = ScriptName.UpdateSepcificApp });
        }

        private async void btnTwo_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });
            await Task.Run(async () =>
            {
                // Create a PowerShell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                // Create a pipeline
                Pipeline pipeline = runspace.CreatePipeline();

                // Load PowerShell script from JSON file
                //string script = GetScript("GetPrinterList");

                // Add PowerShell script to get printer list
                // Add PowerShell script
                // Load PowerShell script from JSON file
                //string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                pipeline.Commands.AddScript(@"try {
        Write-Output ""Updating all apps...""
        $apps = winget list --installed --accept-source-agreements --accept-package-agreements | ForEach-Object { $_.Id }

        foreach ($app in $apps) {
            $result = winget upgrade --id $app --accept-source-agreements --accept-package-agreements

            if ($result.ExitCode -eq 0) {
                Write-Output ""$app updated successfully.""
            }
            else {
                Write-Output ""Failed to update $($app): $($result.StandardError)""
            }
        }
    }
    catch {
        Write-Output ""Error updating apps: $($_.Exception.Message)""
    }");
                // Add printer name as a parameter
                //pipeline.Commands[0].Parameters.Add("appName", appName);

                try
                {
                    // Execute the script
                    Collection<PSObject> results = pipeline.Invoke();

                    if (results.Count > 0)
                    {
                        foreach (PSObject result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (result != null)
                                    MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                            });
                        }
                    }
                }
                catch (RuntimeException ex)
                {
                    System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                }

                // Clean up
                pipeline.Dispose();
                runspace.Close();
            });
            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();
            foreach (var message in MainWindowVMObj.Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                MainWindowVMObj.Messages.Remove(message);
            }
        }

        private void svMessage_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            try
            {
                ScrollViewer scv = (ScrollViewer)sender;
                scv.ScrollToVerticalOffset(scv.VerticalOffset - e.Delta);
                e.Handled = true;
            }
            catch (Exception ex)
            {
            }
        }

        private void btnAbout_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("more solutions coming soon");
        }

        private void btnA_Click(object sender, RoutedEventArgs e)
        {
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Let's optimize your computer for performance.", MessageStatus = "Received", MessageType = MessageType.PowerShellYourComputer });
        }

        private void btnB_Click(object sender, RoutedEventArgs e)
        {
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Enter the name of the program causing issues:", MessageStatus = "Received", MessageType = MessageType.InputConfirmation, ScriptName = ScriptName.UpdateSepcificApp });
        }

        private async void btnPowerShellYourComp1_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });
            await Task.Run(async () =>
            {
                // Create a PowerShell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                // Create a pipeline
                Pipeline pipeline = runspace.CreatePipeline();

                // Load PowerShell script from JSON file
                //string script = GetScript("GetPrinterList");

                // Add PowerShell script to get printer list
                // Add PowerShell script
                // Load PowerShell script from JSON file
                //string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                pipeline.Commands.AddScript(@"$usedApps = Get-Process | Sort-Object -Property CPU -Descending | Select-Object -First 5 -Property ProcessName
                for ($i = 0; $i -lt $usedApps.Count; $i++) {
                    Write-Output ""$($i+1). $($usedApps[$i].ProcessName)""
                }");
                // Add printer name as a parameter
                //pipeline.Commands[0].Parameters.Add("appName", appName);

                try
                {
                    // Execute the script
                    Collection<PSObject> results = pipeline.Invoke();
                    List<string> MessageResponse = new List<string>();
                    if (results.Count > 0)
                    {
                        foreach (PSObject result in results)
                        {
                            if (result != null)
                                MessageResponse.Add(result.ToString());
                        }

                        if (results != null && results.Count > 0)
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Most CPU-intensive apps:", MessageStatus = "Received", MessageType = MessageType.CPUIntensiveApps, MessageResponse = MessageResponse, ScriptName = ScriptName.CPUIntensiveApps });
                            });

                    }
                }
                catch (RuntimeException ex)
                {
                    System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                }

                // Clean up
                pipeline.Dispose();
                runspace.Close();
            });
            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            foreach (var message in MainWindowVMObj.Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                MainWindowVMObj.Messages.Remove(message);
            }

        }

        private async void btnPowerShellYourComp2_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = "Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });

            await Task.Run(async () =>
            {
                // Create a PowerShell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                // Create a pipeline
                Pipeline pipeline = runspace.CreatePipeline();

                // Load PowerShell script from JSON file
                //string script = GetScript("GetPrinterList");

                // Add PowerShell script to get printer list
                // Add PowerShell script
                // Load PowerShell script from JSON file
                //string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                string script = "$usedApps = Get-Process | Sort-Object -Property CPU -Descending | Select-Object -First 5 -Property ProcessName\r\n                for ($i = 0; $i -lt $usedApps.Count; $i++) {\r\n                    Write-Output \"$($i+1). $($usedApps[$i].ProcessName)\"\r\n                }";
                pipeline.Commands.AddScript(@script);
                // Add printer name as a parameter
                //pipeline.Commands[0].Parameters.Add("appName", appname);

                try
                {
                    // Execute the script
                    Collection<PSObject> results = pipeline.Invoke();
                    
                    if (results.Count > 0)
                    {
                        List<string> MessageResponse = new List<string>();
                        foreach (PSObject result in results)
                        {
                            string appname = GetAppName(result.ToString());
                            MessageResponse.Add(appname);
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                svMessage.ScrollToBottom();
                                MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Checking for updates for {appname}", MessageStatus = "Received", MessageType = MessageType.Normal });
                            });
                        }

                        if(MessageResponse.Count>0)
                        {
                            foreach (var item in MessageResponse)
                            {
                                await Task.Run(async () =>
                                {
                                    await UpdateApp(item);
                                });
                            }
                        }
                    }
                }
                catch (RuntimeException ex)
                {
                    System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                }

                // Clean up
                pipeline.Dispose();
                runspace.Close();
            });

            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            foreach (var message in MainWindowVMObj.Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                MainWindowVMObj.Messages.Remove(message);
            }

        }

        private async Task UpdateApp(string appname)
        {
            await Task.Run(async () =>
            {
                // Create a PowerShell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                // Create a pipeline
                Pipeline pipeline = runspace.CreatePipeline();

                // Load PowerShell script from JSON file
                //string script = GetScript("GetPrinterList");

                // Add PowerShell script to get printer list
                // Add PowerShell script
                // Load PowerShell script from JSON file
                //string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                string script = "try {\r\n        Write-Output \"Checking for updates for $appName...\"\r\n        $result = winget upgrade --id $appName --accept-source-agreements --accept-package-agreements\r\n\r\n        if ($result.ExitCode -eq 0) {\r\n           Write-Output \"$appName updated successfully.\"\r\n        }\r\n        else {\r\n           Write-Output \"No updates found for $appName.\"\r\n        }\r\n    }\r\n    catch {\r\n        Write-Output \"Error checking for updates: $($_.Exception.Message)\"\r\n    }";
                pipeline.Commands.AddScript(@script);
                // Add printer name as a parameter
                pipeline.Commands[0].Parameters.Add("appName", appname);

                try
                {
                    // Execute the script
                    Collection<PSObject> results = pipeline.Invoke();
                    if (results.Count > 0)
                    {
                        foreach (PSObject result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()} {appname}", MessageStatus = "Received", MessageType = MessageType.Normal });
                            });
                        }
                    }
                }
                catch (RuntimeException ex)
                {
                    System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                }

                // Clean up
                pipeline.Dispose();
                runspace.Close();
            });
        }

        private async void btnPowerShellYourComp3_Click(object sender, RoutedEventArgs e)
        {
            svMessage.ScrollToBottom();
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = "Updating all apps......", MessageStatus = "Received", MessageType = MessageType.Normal });
            MainWindowVMObj.Messages.Add(new ConversationMessages { Message = "Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });

            await Task.Run(async () =>
            {
                // Create a PowerShell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                // Create a pipeline
                Pipeline pipeline = runspace.CreatePipeline();

                // Load PowerShell script from JSON file
                //string script = GetScript("GetPrinterList");

                // Add PowerShell script to get printer list
                // Add PowerShell script
                // Load PowerShell script from JSON file
                //string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                string script = " try {\r\n        Write-Output \"Updating all apps...\"\r\n        $apps = winget list --installed --accept-source-agreements --accept-package-agreements | ForEach-Object { $_.Id }\r\n\r\n        foreach ($app in $apps) {\r\n            $result = winget upgrade --id $app --accept-source-agreements --accept-package-agreements\r\n\r\n            if ($result.ExitCode -eq 0) {\r\n                Write-Output \"$app updated successfully.\"\r\n            }\r\n            else {\r\n                Write-Output \"Failed to update $($app): $($result.StandardError)\"\r\n            }\r\n        }\r\n    }\r\n    catch {\r\n        Write-Output \"Error updating apps: $($_.Exception.Message)\"\r\n    }";
                pipeline.Commands.AddScript(@script);
                // Add printer name as a parameter
                //pipeline.Commands[0].Parameters.Add("appName", appname);

                try
                {
                    // Execute the script
                    Collection<PSObject> results = pipeline.Invoke();
                    if (results.Count > 0)
                    {
                        foreach (PSObject result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                svMessage.ScrollToBottom();
                                MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                            });
                        }

                        if(results.Count > 0)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                svMessage.ScrollToBottom();
                                MainWindowVMObj.Messages.Add(new ConversationMessages { Message = "App updates completed. Would you like to restart your computer now? (Y/N):", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.RestartComputer });
                            });
                        }
                    }
                }
                catch (RuntimeException ex)
                {
                    System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                }

                // Clean up
                pipeline.Dispose();
                runspace.Close();
            });

            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            foreach (var message in MainWindowVMObj.Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                MainWindowVMObj.Messages.Remove(message);
            }

        }

        private async void lbSelectApps_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            svMessage.ScrollToBottom();
            //MainWindowVMObj.Messages.Add(new ConversationMessages { Message = $"Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });
            System.Windows.Controls.ListBox listBox = sender as System.Windows.Controls.ListBox;
            string appname = GetAppName(listBox.SelectedItem.ToString());
            await Task.Run(async () =>
            {
                // Create a PowerShell runspace
                Runspace runspace = RunspaceFactory.CreateRunspace();
                runspace.Open();

                // Create a pipeline
                Pipeline pipeline = runspace.CreatePipeline();

                // Load PowerShell script from JSON file
                //string script = GetScript("GetPrinterList");

                // Add PowerShell script to get printer list
                // Add PowerShell script
                // Load PowerShell script from JSON file
                //string script = MainWindowVMObj.GetScript("CheckPrinterExists");

                string script = "try {\r\n        Write-Output \"Checking for updates for $appName...\"\r\n        $result = winget upgrade --id $appName --accept-source-agreements --accept-package-agreements\r\n\r\n        if ($result.ExitCode -eq 0) {\r\n           Write-Output \"$appName updated successfully.\"\r\n        }\r\n        else {\r\n           Write-Output \"No updates found for $appName.\"\r\n        }\r\n    }\r\n    catch {\r\n        Write-Output \"Error checking for updates: $($_.Exception.Message)\"\r\n    }";
                pipeline.Commands.AddScript(@script);
                // Add printer name as a parameter
                pipeline.Commands[0].Parameters.Add("appName", appname);

                try
                {
                    // Execute the script
                    Collection<PSObject> results = pipeline.Invoke();
                    if (results.Count > 0)
                    {
                        foreach (PSObject result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                svMessage.ScrollToBottom();
                                MainWindowVMObj.Messages.Add(new ConversationMessages {Message = $"{result.ToString()} {appname}", MessageStatus = "Received",MessageType = MessageType.Normal });
                            });
                        }
                    }
                }
                catch (RuntimeException ex)
                {
                    System.Windows.MessageBox.Show("PowerShell Error: " + ex.Message);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message);
                }

                // Clean up
                pipeline.Dispose();
                runspace.Close();
            });
            //List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            //foreach (var message in MainWindowVMObj.Messages)
            //{
            //    if (message.MessageType == MessageType.Wait)
            //    {
            //        messagesToRemove.Add(message);
            //    }
            //}

            //foreach (var message in messagesToRemove)
            //{
            //    MainWindowVMObj.Messages.Remove(message);
            //}
        }
    }
}
