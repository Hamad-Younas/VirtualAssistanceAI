using Newtonsoft.Json;
using System.Collections.ObjectModel;
using VirtualAssistanceAI.Helpers;
using VirtualAssistanceAI.Models;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Windows.Forms;
using System.Linq;
using OpenAI_API.Moderation;

namespace VirtualAssistanceAI.ViewModels
{
    public class MainWindowVM : BaseHandler
    {
        public MainWindow MainWindowObj { get; set; }
        private ObservableCollection<ConversationMessages> messages;
        public ObservableCollection<ConversationMessages> Messages
        {
            get { return messages; }
            set { messages = value;
                NotifyPropertyChanged("Messages"); 
            }
        }

        #region Commands
        public RelayCommand BtnSendCommand { get; set; }
        public RelayCommand BtnSoundIssueCommand { get; set; }
        public RelayCommand BtnPrinterIssueCommand { get; set; }

        public RelayCommand BtnPasswordHelperCommand { get; set; }
        public RelayCommand BtnWifiIssueCommand { get; set; }

        public RelayCommand AddBlockCommand { get; set; }
        public RelayCommand SetBrightnessCommand { get; set; } 
        public RelayCommand PowerShellOptimizeCommand { get; set; } 
        #endregion

        public MainWindowVM(MainWindow mainWindow)
        {
            MainWindowObj = mainWindow; 
            Messages = new ObservableCollection<ConversationMessages>();
            BtnSendCommand = new RelayCommand(o => OnBtnSendClick(o));
            BtnSoundIssueCommand = new RelayCommand(o=>OnBtnSoundClick(o));
            BtnPrinterIssueCommand = new RelayCommand(o=>OnBtnPrinterClick(o));
            BtnWifiIssueCommand = new RelayCommand(o => OnWifiIssueClick(o));
            BtnPasswordHelperCommand = new RelayCommand(o => OnPasswordHelperClick(o));
            AddBlockCommand = new RelayCommand(o => OnAddBlcok(o));
            SetBrightnessCommand = new RelayCommand(o => OnSetBrightness(o));
            PowerShellOptimizeCommand = new RelayCommand(o => OnPowerShellOptimize(o));
            // Load PowerShell scripts from JSON file
            scriptSettingsList = LoadPowerShellScriptFromJson();
            //AddBlockScript();
        }

        private void OnPowerShellOptimize(object o)
        {
            RunPowerShellOptimizeScript();
        }

        private async void OnSetBrightness(object o)
        {
            MainWindowObj.svMessage.ScrollToBottom();
            MainWindowObj.btnScrollUpDown.IsChecked = true;
            Messages.Add(new ConversationMessages { Message = $"Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });

            await Task.Run(async () =>
            {
                await SetBrightness();
            });

            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            foreach (var message in Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                Messages.Remove(message);
            }
        }

        private async void OnAddBlcok(object o)
        {
            MainWindowObj.svMessage.ScrollToBottom();
            MainWindowObj.btnScrollUpDown.IsChecked = true;
            Messages.Add(new ConversationMessages { Message = $"Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });


            await Task.Run(async () =>
            {
                await AddBlockScript();
            });

            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            foreach (var message in Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                Messages.Remove(message);
            }
        }

        public async void OnPasswordHelperClick(object o)
        {
            MainWindowObj.svMessage.ScrollToBottom();
            MainWindowObj.btnScrollUpDown.IsChecked = true;
            Messages.Add(new ConversationMessages { Message = $"Would you like to share your current Wi-Fi password? (y/n)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.AskToUserWifiShare });
        }

        public async void OnWifiIssueClick(object o)
        {
            MainWindowObj.svMessage.ScrollToBottom();
            MainWindowObj.btnScrollUpDown.IsChecked = true;
            Messages.Add(new ConversationMessages { Message = $"Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });
            WebsiteWifiTroubleshooting2();

            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            foreach (var message in Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                Messages.Remove(message);
            }
            
        }

        private async Task WebsiteWifiTroubleshooting2()
        {
            Messages.Add(new ConversationMessages { Message = $"Please enter the website URL or gateway IP address (e.g., google.com or 192.168.0.1):", MessageStatus = "Received", MessageType = MessageType.InputConfirmation,ScriptName= ScriptName.UrlValidation });
           
        }

        private async void OnBtnSoundClick(object o)
        {
            MainWindowObj.svMessage.ScrollToBottom();
            MainWindowObj.btnScrollUpDown.IsChecked = true;
            Messages.Add(new ConversationMessages { Message = "Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });
            await Task.Run(async () =>
            {
                await ResolveSoundIssue();
            });

            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            foreach (var message in Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                Messages.Remove(message);
            }
        }

        private async void OnBtnPrinterClick(object o)
        {
            MainWindowObj.svMessage.ScrollToBottom();
            MainWindowObj.btnScrollUpDown.IsChecked = true;
            Messages.Add(new ConversationMessages { Message = $"Please wait...", MessageStatus = "Received", MessageType = MessageType.Wait });
            
            await Task.Run(async () =>
            {
                await ResolvePrinterIssue();
            });

            List<ConversationMessages> messagesToRemove = new List<ConversationMessages>();

            foreach (var message in Messages)
            {
                if (message.MessageType == MessageType.Wait)
                {
                    messagesToRemove.Add(message);
                }
            }

            foreach (var message in messagesToRemove)
            {
                Messages.Remove(message);
            }

        }

        private async void OnBtnSendClick(object o)
        {
            string topic = o.ToString();

            //if (string.IsNullOrEmpty(AIHelper.OpenAIKey)) {
            //    System.Windows.MessageBox.Show("Please Enter your API key");
            //    return;
            //}

            Messages.Add(new ConversationMessages() { Message = topic, MessageStatus = "Sent", MessageType = MessageType.Normal });
            var response = await GetResponse(topic);
            if (!response.isIssue)
            {
                await AIHelper.AiHelperInstance.GetAIResponse(topic, Messages, MainWindowObj);
            }
            else
            {
                if (response.userInput.ToLower().Contains("wifi issue") || response.userInput.ToLower().Contains("website wifi troubleshooting") || response.userInput.ToLower().Contains("websitewifitroubleshooting") ||
                    response.userInput.ToLower().Contains("wifitroubleshooting") || response.userInput.ToLower().Contains("internet") || response.userInput.ToLower().Contains("wifi") || response.userInput.ToLower().Contains("wif-fi") || response.userInput.ToLower().Contains("chrome") || response.userInput.ToLower().Contains("browser"))
                {
                    WebsiteWifiTroubleshooting();
                    //Messages.Add(new ConversationMessages { Message = $"Please wait I am working on wifi issue", MessageStatus = "Received", MessageType = MessageType.Normal });
                }
                else if (response.userInput.ToLower().Contains("sound issue") || response.userInput.ToLower().Contains("sound") || response.userInput.ToLower().Contains("hear") || response.userInput.ToLower().Contains("audio"))
                {
                    await Task.Run(async () =>
                    {
                        await ResolveSoundIssue();
                    });
                    //Messages.Add(new ConversationMessages { Message = $"Please wait I am working on sound issue", MessageStatus = "Received", MessageType = MessageType.Normal });
                }
                else if (response.userInput.ToLower().Contains("printer issue") || response.userInput.ToLower().Contains("print") || response.userInput.ToLower().Contains("printing") || response.userInput.ToLower().Contains("printer") || response.userInput.ToLower().Contains("ricoh") || response.userInput.ToLower().Contains("prints") || response.userInput.ToLower().Contains("inkjet"))
                {
                    await Task.Run(async () =>
                    {
                        await ResolvePrinterIssue();
                    });
                    //Messages.Add(new ConversationMessages { Message = $"Please wait I am working on printer issue", MessageStatus = "Received" });
                    //await ResolvePrinterIssue();
                }
                else if (response.userInput.ToLower().Contains("password issue") || response.userInput.ToLower().Contains("wifi password") || response.userInput.ToLower().Contains("pasword") || response.userInput.ToLower().Contains("passwrod") || response.userInput.ToLower().Contains("password"))
                {
                    //Messages.Add(new ConversationMessages { Message =$"Please wait I am working on password issue", MessageStatus = "Received", MessageType = MessageType.Normal });
                    Messages.Add(new ConversationMessages { Message = $"Would you like to share your current Wi-Fi password? (y/n)", MessageStatus = "Received", MessageType = MessageType.Confirmation,ScriptName = ScriptName.AskToUserWifiShare });
                    //PasswordResetAssistance();
                }
                else if (response.userInput.ToLower().Contains("set brightness") || response.userInput.ToLower().Contains("brightness level")
                    || response.userInput.ToLower().Contains("brightness level set") || response.userInput.ToLower().Contains("bright") || response.userInput.ToLower().Contains("brighter") || response.userInput.ToLower().Contains("brightness"))
                {
                    Task.Run(async () =>
                    {
                        //Messages.Add(new ConversationMessages { Message = $"Please wait I am working on password issue", MessageStatus = "Received", MessageType = MessageType.Normal });
                        await SetBrightness();
                    });
                    
                }
                else if (response.userInput.ToLower().Contains("powershell script optimize") || response.userInput.ToLower().Contains("optimize") || response.userInput.ToLower().Contains("update") || response.userInput.ToLower().Contains("secure"))
                {
                    RunPowerShellOptimizeScript();
                }
            }
          
        }



        #region Scripts
        private void RunPowerShellOptimizeScript()
        {
            Messages.Add(new ConversationMessages { Message = $"Oh no, it sounds like you're having an issue with your computer running slow.", MessageStatus = "Received", MessageType = MessageType.PowerShellOptimizieChoice });
        }

        private void RunPowerShellOptimizeScriptOld()
        {
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                using (Pipeline pipeline = runspace.CreatePipeline())
                {
                    pipeline.Commands.AddScript(@"# Run a diagnostic to determine the most used apps
$usedApps = Get-Process | Sort-Object -Property CPU -Descending | Select-Object -First 5 -Property ProcessName

for ($i = 0; $i -lt $usedApps.Count; $i++) {
    Write-Output ""$($i+1). $($usedApps[$i].ProcessName)""
}");


                    try
                    {
                        var results = pipeline.Invoke();
                        List<string> MostusedAppList = new List<string>();
                        foreach (var result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (result != null)
                                    MostusedAppList.Add(result.ToString());
                            });
                        }

                        if (MostusedAppList.Count > 0)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                Messages.Add(new ConversationMessages { Message = $"Updating device and current apps causing issues for optimal performance..", MessageStatus = "Received", MessageType = MessageType.RunpowerShell, MessageResponse = MostusedAppList });
                            });

                        }
                    }
                    catch (RuntimeException ex)
                    {
                        MessageBox.Show("PowerShell Error: " + ex.Message);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }
        }

        public async Task AddBlockScript()
        {
            // Import the Chrome WebDriver
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                using (Pipeline pipeline = runspace.CreatePipeline())
                {
                    pipeline.Commands.AddScript(@"# Function to get Chrome installation path from registry
function Get-ChromeInstallationPath {
    $chromePath = """"

    # Paths to check for Chrome's installation path
    $regPaths = @(
        ""HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"",
        ""HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"",
        ""HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe""
    )

    foreach ($regPath in $regPaths) {
        try {
            $path = Get-ItemProperty -Path $regPath -Name ""(default)"" -ErrorAction Stop
            if ($path.""(default)"") {
                $chromePath = $path.""(default)""
                break
            }
        } catch {
            # Do nothing, try the next path
        }
    }

    if ($chromePath) {
        return [System.IO.Path]::GetDirectoryName($chromePath)
    } else {
        Write-Output ""Google Chrome is not installed.""
    }
}

$ChromeInstallPath=""""

# Get and display the Chrome installation path
$chromeInstallationPath = Get-ChromeInstallationPath
if ($chromeInstallationPath) {
    $ChromeInstallPath = $chromeInstallationPath
}



# Create a new WScript.Shell object for sending keyboard input
$wshell = New-Object -ComObject WScript.Shell

# Wait for 5 seconds before proceeding
Start-Sleep -Seconds 5

# Check if the Chrome installation path exists
if (Test-Path $ChromeInstallPath) {
    try {
        # Start Chrome and navigate to the web store page
        Start-Process -FilePath ""$ChromeInstallPath\chrome.exe"" -ArgumentList ""https://chrome.google.com/webstore/detail/gighmmpiobklfepjocnamgkkbiglidom"" -WindowStyle Maximized -ErrorAction Stop

        # Wait for Chrome to load
        Start-Sleep -Seconds 10

        # Search for ""Add to Chrome"" using Ctrl+F
        $wshell.AppActivate(""Google Chrome"")
        Start-Sleep -Seconds 2
        $wshell.SendKeys(""^f"")
        Start-Sleep -Seconds 2
        $wshell.SendKeys(""Add to Chrome"")
        Start-Sleep -Seconds 2
        $wshell.SendKeys(""{ENTER}"")

        # Navigate through the installation prompts
        Start-Sleep -Seconds 2
        $wshell.SendKeys(""{TAB}"")
        Start-Sleep -Seconds 2
        $wshell.SendKeys(""{TAB}"")
        Start-Sleep -Seconds 2
        $wshell.SendKeys(""{TAB}"")
        Start-Sleep -Seconds 2
        $wshell.SendKeys("" "")
        Start-Sleep -Seconds 5
        $wshell.SendKeys(""{ENTER}"")
        Start-Sleep -Seconds 10
        $wshell.SendKeys(""{TAB}"")
        Start-Sleep -Seconds 5
        $wshell.SendKeys(""{ENTER}"")
    }
    catch {
        Write-Output ""An error occurred while installing the Chrome extension: $($_.Exception.Message)""
    }
}
else {
    Write-Output ""Couldn't find the Chrome installation path $ChromeInstallPath. Please check the path and update it in the script.""
}");


                    try
                    {
                        var results = pipeline.Invoke();
                        foreach (var result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (result != null)
                                {
                                    if (result.ToString() == "True" || result.ToString() == "False")
                                    {
                                        if (bool.Parse(result.ToString()))
                                        {
                                            Messages.Add(new ConversationMessages { Message = $"Ad blocker installed successfully.", MessageStatus = "Received", MessageType = MessageType.Normal });
                                        }
                                    }
                                } 
                            });
                        }
                    }
                    catch (RuntimeException ex)
                    {
                        MessageBox.Show("PowerShell Error: " + ex.Message);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }
        }

        public void AskToUserWifiShare(ScriptName scriptName)
        {
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                using (Pipeline pipeline = runspace.CreatePipeline())
                {
                    pipeline.Commands.AddScript(@"# Function to retrieve active window URL
function Get-ActiveWindowURL {
    try {
        $signature = @""
[DllImport(""user32.dll"", SetLastError=true, CharSet=CharSet.Auto)]
public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
[DllImport(""user32.dll"", SetLastError=true, CharSet=CharSet.Auto)]
public static extern IntPtr GetForegroundWindow();
""@

        $type = Add-Type -MemberDefinition $signature -Name GetActiveWindowURL -Namespace User32 -PassThru
        $hwnd = $type::GetForegroundWindow()
        $sb = New-Object System.Text.StringBuilder(256)
        $type::GetWindowText($hwnd, $sb, $sb.Capacity) | Out-Null
        return $sb.ToString()
    }
    catch {
        Write-Output ""Error retrieving active window URL: $($_.Exception.Message)"" -ForegroundColor Red
        return $null
    }
}

# Function to retrieve Wi-Fi security key
function Get-WiFiSecurityKey {
    $profileName = (netsh wlan show interfaces | Select-String ""SSID"").Line.Split(':')[1].Trim()
    $securityKey = (netsh wlan show profile name=""$profileName"" key=clear | Select-String ""Key Content"").Line.Split(':')[1].Trim()
    return $securityKey
}

# Ask the user if they want to share their Wi-Fi password
$wifiSSID = (netsh wlan show interfaces | Select-String ""SSID"").Line.Split(':')[1].Trim()
    if ($wifiSSID) {
        Write-Output ""Your current Wi-Fi Network is: $wifiSSID""
        $wifiSecurityKey = Get-WiFiSecurityKey
        if ($wifiSecurityKey) {
            Write-Output ""Your current Wi-Fi password is: $wifiSecurityKey""
        } else {
            Write-Output ""No Wi-Fi security key found.""
        }
}
");

                    // Add printer name as a parameter
                    //pipeline.Commands[0].Parameters.Add("shareWifiPassword", "yes");

                    try
                    {
                        var results = pipeline.Invoke();
                        foreach (var result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (result != null)
                                    Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                            });
                        }
                        if (results != null && results.Count > 0)
                        {
                            Messages.Add(new ConversationMessages { Message = $"Are you having any other password issues that I can help with? (Y/N):", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.PasswordManager });
                        }

                    }
                    catch (RuntimeException ex)
                    {
                        MessageBox.Show("PowerShell Error: " + ex.Message);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }
        }

        private async Task WebsiteWifiTroubleshooting()
        {
            // Import the Chrome WebDriver
            using (Runspace runspace = RunspaceFactory.CreateRunspace())
            {
                runspace.Open();

                using (Pipeline pipeline = runspace.CreatePipeline())
                {
                    pipeline.Commands.AddScript(@"
                            try {
                                $activeUrl = (Get-Process -Id (Get-Process -Id $pid).MainWindowHandle).MainWindowTitle
                            } catch {
                                Write-Output 'Error: Unable to retrieve the active window URL. Please make sure you are running the script in an interactive session.' -ForegroundColor Red
                                exit
                            }

                            Write-Output 'Testing Internet connection...'
                            $pingResult = Test-NetConnection -ComputerName google.com

                            if ($pingResult.PingSucceeded) {
                                Write-Output 'Pinging $activeUrl...'
                                $pingUrl = Test-NetConnection -ComputerName $activeUrl.TrimStart('https://', 'http://')

                                if ($pingUrl.PingSucceeded) {
                                    Write-Output 'Clearing browser cache and restarting the browser...'
                                    $browsers = Get-Process -Name '*chrome*', '*firefox*', '*edge*', '*iexplore*'
                                    foreach ($browser in $browsers) {
                                        $browser.CloseMainWindow() | Out-Null
                                    }

                                    Write-Output 'Testing $activeUrl again...'
                                    $pingUrl = Test-NetConnection -ComputerName $activeUrl.TrimStart('https://', 'http://')

                                    if ($pingUrl.PingSucceeded) {
                                        Write-Output 'The issue seems to be resolved after clearing the cache and restarting the browser.'
                                    }
                                    else {
                                        Write-Output 'The issue persists. Trying further troubleshooting steps...'
                                    }
                                }
                                else {
                                    Write-Output 'The issue appears to be with the specific website $activeUrl. Trying to test it in a different browser.'
                                    $defaultBrowser = (Get-ItemPropertyValue 'HKCU:\\Software\\Microsoft\\Windows\\Shell\\Associations\\UrlAssociations\\http\\UserChoice' -Name 'ProgId')
                                    switch ($defaultBrowser) {
                                        'ChromeHTML' { $browserPath = 'chrome.exe' }
                                        'FirefoxURL' { $browserPath = 'firefox.exe' }
                                        'AppXq0fevzme2pys62n3e0fbqa7peapykr8v' { $browserPath = 'msedge.exe' } # Edge
                                        'IE.HTTP' { $browserPath = 'iexplore.exe' }
                                        default { Write-Output 'Unknown browser.' }
                                    }
                                    
                                    if ($browserPath) {
                                        Start-Process $browserPath -ArgumentList $activeUrl
                                    } else {
                                        Write-Output 'Unable to determine the default browser. Please open the URL manually in a browser.'
                                    }

                                    Write-Output 'Testing $activeUrl again in a different browser...'
                                    $pingUrl = Test-NetConnection -ComputerName $activeUrl.TrimStart('https://', 'http://')

                                    if ($pingUrl.PingSucceeded) {
                                        Write-Output 'The issue seems to be resolved after testing it in a different browser.'
                                    }
                                    else {
                                        Write-Output 'Clearing browser cache and restarting the browser...'
                                        $browsers = Get-Process -Name '*chrome*', '*firefox*', '*edge*', '*iexplore*'
                                        foreach ($browser in $browsers) {
                                            $browser.CloseMainWindow() | Out-Null
                                        }

                                        Write-Output 'Testing $activeUrl again...'
                                        $pingUrl = Test-NetConnection -ComputerName $activeUrl.TrimStart('https://', 'http://')

                                        if ($pingUrl.PingSucceeded) {
                                            Write-Output 'The issue seems to be resolved after clearing the cache and restarting the browser.'
                                        }
                                        else {
                                            Write-Output 'The issue persists. Further troubleshooting is required.'
                                        }
                                    }
                                }
                            }
                            else {
                                Write-Output 'Restarting the WiFi...'
                                $netAdapter = Get-NetAdapter -Name '*Wi-Fi*'
                                $netAdapter | Disable-NetAdapter -Confirm:$false
                                Start-Sleep -Seconds 5
                                $netAdapter | Enable-NetAdapter -Confirm:$false

                                Start-Sleep -Seconds 30

                                Write-Output 'Testing internet speed...'
                                $speedTestUrl = 'https://www.speedtest.net/'
                                $speedTestResult = Test-NetConnection -ComputerName $speedTestUrl.TrimStart('https://', 'http://')

                                if ($speedTestResult.PingSucceeded) {
                                    Write-Output 'The issue seems to be resolved. Opening $activeUrl...'
                                    Start-Process ($activeUrl.Split('/')[0] + '//' + $activeUrl.Split('/')[2])
                                }
                                else {
                                    Write-Output 'Internet speed is too slow or not working. Restarting the router...'
                                    $routerLoginInfo = @('admin/admin', 'admin/Admin', 'admin/password', 'admin/1234') # Router login credentials
                                    $routerIP = '192.168.0.1' # Router IP

                                    foreach ($credential in $routerLoginInfo) {
                                        $username = $credential.Split('/')[0]
                                        $password = $credential.Split('/')[1]
                                        Write-Output 'Trying to restart the router using the IP admin console with the credential: $username/$password...'
                                        $url = 'http://$routerIP/'
                                        $basicAuthValue = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes('$username:$password'))
                                        $headers = @{ Authorization = 'Basic $basicAuthValue' }
                                        try {
                                            $response = Invoke-WebRequest -Uri $url -Headers $headers -Method GET
                                            if ($response.StatusCode -eq 200) {
                                                Write-Output 'Router restarted successfully.'
                                            }
                                            else {
                                                Write-Output 'Failed to authenticate to the router using the provided credentials.'
                                            }
                                        }
                                        catch {
                                            Write-Output 'An error occurred while trying to access the router: $($_.Exception.Message)'
                                        }

                                        Start-Sleep -Seconds 300 # Wait for 5 minutes

                                        Write-Output 'Testing internet connection again...'
                                        $pingResult = Test-NetConnection -ComputerName google.com

                                        if ($pingResult.PingSucceeded) {
                                            Write-Output 'The issue seems to be resolved after restarting the router.'
                                            break
                                        }
                                        else {
                                            Write-Output 'The issue persists. Please unplug your router, wait for 60 seconds, plug it back in, and wait for 5-10 minutes for it to power up and test the internet again.'
                                        }
                                    }
                                }
                            }
                        ");


                    try
                    {
                        var results = pipeline.Invoke();
                        foreach (var result in results)
                        {
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (result != null)
                                    Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                            });
                        }
                    }
                    catch (RuntimeException ex)
                    {
                        MessageBox.Show("PowerShell Error: " + ex.Message);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }
        }

        private async Task ResolveSoundIssue()
        {
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            //string script = GetScript("CheckMaximizeDeviceVolume");

            // Add PowerShell script to get printer list
            pipeline.Commands.AddScript(@"# Function to check if a browser window is muted
function IsBrowserMuted($browserProcess) {
    $browserHWND = $browserProcess.MainWindowHandle
    $muteStatus = (New-Object -ComObject Shell.Application).Windows() | Where-Object { $_.HWND -eq $browserHWND } | ForEach-Object { $_.Document } | ForEach-Object { $_.getElementsByClassName('mute-button')[0].getAttribute('aria-pressed') }
    return $muteStatus -eq 'true'
}

# Function to play continuous beep sound
function PlayContinuousBeep {
    for ($i = 1; $i -le 10; $i++) {
        [console]::Beep(1000, 200)
        Start-Sleep -Milliseconds 200
    }
}

# Check and maximize the current browser's volume levels, make sure nothing is muted
Write-Output ""Checking and maximizing the current browser's volume levels...""
$browsers = Get-Process -Name ""*chrome*"", ""*firefox*"", ""*edge*"", ""*iexplore*""
$processedBrowsers = @()
foreach ($browser in $browsers) {
    try {
        $browserName = $browser.ProcessName
        if ($processedBrowsers -notcontains $browserName) {
            if (IsBrowserMuted $browser) {
                Write-Output ""The $browserName browser is muted. Please unmute it for sound.""
            } else {
                Write-Output ""The $browserName browser is not muted. Maximizing volume...""
                $processedBrowsers += $browserName
                $wshShell = New-Object -ComObject WScript.Shell
                for ($i = 0; $i -lt 50; $i++) {
                    $wshShell.SendKeys([char]175)
                    Start-Sleep -Milliseconds 50
                }
                Write-Output ""Volume maximized for $browserName""
            }
        }
    }
    catch {
        Write-Output ""Error occurred while trying to control volume for $($browser.ProcessName): $($_.Exception.Message)""
    }
}

# Check and maximize the device's volume levels
Write-Output ""Checking and maximizing the device's volume levels...""
try {
    # Set volume to maximum
    (New-Object -ComObject Shell.Application).ToggleMute() | Out-Null
} catch {
    Write-Output ""An error occurred while trying to maximize the device's volume: $_""
}

# Play continuous beep sound
Write-Output ""Playing continuous beep sound...""
PlayContinuousBeep");

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();
                foreach (PSObject result in results)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (result != null)
                            Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                    });
                }

                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    Messages.Add(new ConversationMessages { Message = $"Is there still no sound? (Y/N)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.AskStillNoSound });
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

        public async Task SetBrightness()
        {
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            string script = GetScript("DisplayBrightnessLevel");

            // Add PowerShell script to get printer list
            pipeline.Commands.AddScript(@script);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();


                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    foreach (PSObject result in results)
                    {
                        CurrentBrightnessLevel = 100;
                        if (result != null)
                            Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
                    }

                    Messages.Add(new ConversationMessages { Message = $"Would you like to adjust the brightness level? (Y/N)", MessageStatus = "Received", MessageType = MessageType.Confirmation, ScriptName = ScriptName.AskToUserAdjustBrightness });
                });

                //AskToUserAdjustBrightness();
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

        public int CurrentBrightnessLevel { get; set; }
        public void AskToUserAdjustBrightness(int newLevel)
        {
            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            string script = GetScript("AskToUserAdjustBrightness");

            // Add PowerShell script to get printer list
            pipeline.Commands.AddScript(@script);

            // Add printer name as a parameter
            pipeline.Commands[0].Parameters.Add("newLevel", newLevel);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();


                foreach (PSObject result in results)
                {
                    CurrentBrightnessLevel = newLevel;
                    if (result != null)
                        Messages.Add(new ConversationMessages { Message = $"{result.ToString()}", MessageStatus = "Received", MessageType = MessageType.Normal });
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

        List<ScriptSettings> scriptSettingsList;
        public async Task ResolvePrinterIssue()
        {


            // Create a PowerShell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // Create a pipeline
            Pipeline pipeline = runspace.CreatePipeline();

            // Load PowerShell script from JSON file
            //string script = GetScript("GetPrinterList");
            string script = GetScript("GetPrinterList");

            // Add PowerShell script to get printer list
            pipeline.Commands.AddScript(@script);

            try
            {
                // Execute the script
                Collection<PSObject> results = pipeline.Invoke();

                //PrinterList = new List<string>();
                //PrinterList.Add("Select the printer you want to troubleshoot");

                // Store the printer list

                List<string> MessageResponse = new List<string>();
                foreach (PSObject result in results)
                {
                    // Assuming you want to store printer names in a list
                    string printerName = result.Properties["Name"].Value.ToString();
                    // Do whatever you want with the printerName, such as adding it to a list, displaying it, etc.
                    //Console.WriteLine(printerName); // For demonstration, shows printer name in a MessageBox
                    //Messages.Add(new ConversationMessages { Message = $"{printerName}", MessageStatus = "Received" });
                    //PrinterList.Add(printerName);
                    //lbPromt.ItemsSource = printerName;

                    //lbPromt.Items.Add(printerName);

                    MessageResponse.Add(printerName);
                }

                if (MessageResponse.Count > 0)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        Messages.Add(new ConversationMessages { Message = $"Select the printer you want to troubleshoot", MessageStatus = "Received", MessageType = MessageType.Issue, MessageResponse = MessageResponse });
                    });

                }



                //System.Windows.MessageBox.Show("Select the printer you want to troubleshoot");
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

        public string GetScript(string scriptName)
        {
            ScriptSettings scriptSettings = scriptSettingsList.Find(script => script.ScriptName.Equals(scriptName, StringComparison.OrdinalIgnoreCase));
            return scriptSettings?.Script;
        }

        public List<ScriptSettings> LoadPowerShellScriptFromJson()
        {
            string jsonContent = File.ReadAllText("Scripts.json");
            return JsonConvert.DeserializeObject<List<ScriptSettings>>(jsonContent);
        }

        public async Task<(string userInput, bool isIssue)> GetResponse(string userInput)
        {
            bool isIssue = false;

            if (userInput.ToLower().Contains("wifi issue") || userInput.ToLower().Contains("powershell script optimize") ||
                userInput.ToLower().Contains("sound issue") ||
                userInput.ToLower().Contains("printer issue") || userInput.ToLower().Contains("set brightness") ||
                userInput.ToLower().Contains("brightness level") || userInput.ToLower().Contains("brightness level set") ||
                userInput.ToLower().Contains("website wifi troubleshooting") || userInput.ToLower().Contains("websitewifitroubleshooting") ||
                userInput.ToLower().Contains("wifitroubleshooting") || userInput.ToLower().Equals("wifitroubleshooting") ||
                userInput.ToLower().Contains("password issue") || userInput.ToLower().Contains("wifi password"))
            {
                isIssue = true;
            }

            return (userInput, isIssue);
        } 
        #endregion
    }

    public class ScriptSettings
    {
        [JsonProperty("ScriptName")]
        public string ScriptName { get; set; }

        [JsonProperty("script")]
        public string Script { get; set; }
    }
}
