using AutoIt;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace Autoit_CiscoVPN
{
    class AutoLogin : AppConfig
    {
        int waitTime = 10;
        string messageBoxTitle = "CiscoVPN_Autoit";
        Stopwatch stopwatch;

        #region Cisco_MainWindow handles
        string domainWindowTitle = "Cisco AnyConnect Secure Mobility Client";
        string domainTextbox = "[CLASS:Edit; INSTANCE:1]";
        string domainDropdown = "[CLASS:ComboBox; INSTANCE:1]";
        string connectButton = "[CLASS:Button; TEXT:Connect; INSTANCE:1]";
        string disconnectButton = "[CLASS:Button; TEXT:Disconnect; INSTANCE:1]";
        #endregion

        #region Cisco_loginWindow handles
        string loginWindowTitle = "Cisco AnyConnect | " + domainName;
        string groupDropdown = "[CLASS:ComboBox; INSTANCE:1]";
        string usernameTextbox = "[CLASS:Edit; INSTANCE:1]";
        string passwordTextbox = "[CLASS:Edit; INSTANCE:2]";
        string okButton = "[CLASS:Button; TEXT:OK; INSTANCE:1]";
        #endregion

        #region Cisco_TermsandCondition handles
        string termsandConditonWindowTitle = "Cisco AnyConnect";
        string acceptButton = "[CLASS:Button; TEXT:Accept; INSTANCE:1]";
        #endregion

        #region Common methods
        private void ShowErrorPopup(string errorMsg)
        {
            errorMsg += "\n\nScript will STOP now. Verify the error and restart the script.";
            DialogResult dialog = MessageBox.Show(errorMsg, messageBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (dialog == DialogResult.OK)  //Close console window after clicking ok button
                Environment.Exit(0);
        }

        private void WaitForWindow(string windowID, int waitTime = 10)
        {
            AutoItX.AutoItSetOption("WinTitleMatchMode", 3);
            var handle = AutoItX.WinWaitActive(windowID, "", waitTime);

            if(windowID != termsandConditonWindowTitle) //Dont show errorpopup if there is no terms and condition popup
            {
                if (AutoItX.WinExists(windowID) == 0) //Autoit return 1 for success and 0 for failure
                    ShowErrorPopup($"'{windowID}' : unable to locate this window title");
            }
            
        }

        private void CheckControlVisibility(string windowTitle, string controlID)
        {
            var isControlVisible = AutoItX.ControlCommand(windowTitle, "", controlID, "IsVisible", "");

            if(controlID == disconnectButton)
            {
                if (Int16.Parse(isControlVisible) == 1) //Autoit will return 1 if the control is visible
                {
                    AutoItX.ControlClick(domainWindowTitle, "", disconnectButton); //Uncomment this line if you need disconnect vpn if its already connected
                    Environment.Exit(0); 
                    //ShowErrorPopup($"Cisco VPN is ALREADY CONNECTED.");
                }
            }
            else
            {
                if (Int16.Parse(isControlVisible) == 0) //Autoit will return 1 if the control is visible
                    ShowErrorPopup($"CONTROL NOT VISIBLE\nwindowtitle={windowTitle}\ncontrolID={controlID}");
            }
        }

        private void WaitTillControlEnabled(string windowTitle, string controlID)
        {
            stopwatch = new Stopwatch();
            stopwatch.Start();

            //wait till the control enable and also wait upto the waittime declared
            while (Int16.Parse(AutoItX.ControlCommand(windowTitle, "", controlID, "IsEnabled", "")) == 0 && stopwatch.Elapsed.Seconds <= waitTime)
            {
                AutoItX.Sleep(100);
            }
            stopwatch.Stop();
            AutoItX.Sleep(500); //Added this sleep for stability
        }
        #endregion


        public AutoLogin()
        {
            OpenCiscoVPN();
            ConnectToDomain();
            LoginUsingUserCredentials();
            AcceptTermsandConditionPopup(); //Will accept terms & conditions popup ifany
        }



        private void OpenCiscoVPN()
        {
            if (File.Exists(ciscoExePath))
                AutoItX.Run(ciscoExePath, "", 1);
            else
                ShowErrorPopup("'INSTALL CiscoVPN' or 'Enter the correct path in CONFIG file'.");
        }


        private void ConnectToDomain()
        {
            WaitForWindow(domainWindowTitle);

            //Check all elements are visible in particular window
            CheckControlVisibility(domainWindowTitle, disconnectButton);
            CheckControlVisibility(domainWindowTitle, domainTextbox);
            CheckControlVisibility(domainWindowTitle, domainDropdown);

            //Check textbox value and select domain from dropdown
            string defaultDomainVal = AutoItX.ControlGetText(domainWindowTitle, "", domainTextbox);

            if (defaultDomainVal != domainName && string.IsNullOrEmpty(defaultDomainVal) == false)
            {
                AutoItX.ControlCommand(domainWindowTitle, "", domainDropdown, "AddString", domainName);
                AutoItX.ControlCommand(domainWindowTitle, "", domainDropdown, "SelectString", domainName);
            }
            else
                AutoItX.ControlClick(domainWindowTitle, "", connectButton);

            Console.WriteLine("-- Connecting to domain");
        }

        private void LoginUsingUserCredentials()
        {
            WaitForWindow(loginWindowTitle);

            //Check all elements are visible in this particular window
            CheckControlVisibility(loginWindowTitle, usernameTextbox);
            CheckControlVisibility(loginWindowTitle, passwordTextbox);
            CheckControlVisibility(loginWindowTitle, okButton);

            CheckforGroupSelector(loginWindowTitle, groupDropdown);
            WaitTillControlEnabled(loginWindowTitle, okButton);
            AutoItX.ControlSetText(loginWindowTitle, "", usernameTextbox, userName);
            AutoItX.ControlSetText(loginWindowTitle, "", passwordTextbox, passWord);
            AutoItX.ControlClick(loginWindowTitle, "", okButton);
        }

        // Check for terms and conditions popup and click accept button if any
        private void AcceptTermsandConditionPopup()
        {
            WaitForWindow(termsandConditonWindowTitle, 4); // wait for 4 secs
            if(AutoItX.WinExists(termsandConditonWindowTitle) == 1)
            {
                CheckControlVisibility(termsandConditonWindowTitle, acceptButton);

                AutoItX.ControlClick(termsandConditonWindowTitle, "", acceptButton);
            }
        }

        private void CheckforGroupSelector(string windowTitle, string controlID)
        {
            string GroupdropdownHandle = AutoItX.ControlCommand(windowTitle, "", controlID, "IsVisible", "");
            if (Int16.Parse(GroupdropdownHandle) == 1) //If control isenabled, autoit will return 1
            {
                if (!string.IsNullOrEmpty(groupName))
                {
                    AutoItX.ControlCommand(windowTitle, "", groupDropdown, "SelectString", groupName);
                    string selectedGroupName = AutoItX.ControlCommand(windowTitle, "", groupDropdown, "GetCurrentSelection", groupName);


                    if (selectedGroupName != groupName)
                        ShowErrorPopup($"Unable to find 'GROUP: {groupName}'. Check it in CONFIG file.");
                }
                else
                    ShowErrorPopup($"Enter valid GROUP {groupName} inside CONFIG file");
            }
        }
    }
}
