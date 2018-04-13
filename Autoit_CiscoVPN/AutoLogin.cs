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

        #region Cisco_CertificationCheck handles
        string certificationWindowTitle = "Cisco AnyConnect Secure Mobility Client";
        string certificationWindowText = "Security Warning: Untrusted Server Certificate!";
        string keepMeSafeButton = "[CLASS:Button; TEXT:Keep Me Safe; INSTANCE:1]";
        string changeSettingButton = "[CLASS:Button; TEXT:Change Setting...; INSTANCE:2]";
        string connectAnywayButton = "[CLASS:Button; ID:1067; INSTANCE:2]";
        string cancelConnectionButton = "[CLASS:Button; TEXT:Cancel Connection; INSTANCE:1]";
        #endregion

        #region Common methods
        private void ShowErrorPopup(string errorMsg)
        {
            errorMsg += "\n\nScript will STOP now. Verify the error and restart the script.";
            DialogResult dialog = MessageBox.Show(errorMsg, messageBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (dialog == DialogResult.OK)  //Close console window after clicking ok button
                Environment.Exit(0);
        }

        private void WaitForWindow(string windowID, int waitTime = 10, string windowText = "")
        {
            AutoItX.AutoItSetOption("WinTitleMatchMode", 3);
            var handle = AutoItX.WinWaitActive(windowID, windowText, waitTime);

            //if ((windowID != termsandConditonWindowTitle) || (windowID != certificationWindowTitle)) //Dont show errorpopup if there is no terms and condition popup
            if ((windowID != termsandConditonWindowTitle) && (windowID != certificationWindowTitle)) //Dont show errorpopup if there is no terms and condition popup
            {
                //do nothing
                if (AutoItX.WinExists(windowID) == 0) //Autoit return 1 for success and 0 for failure
                    ShowErrorPopup($"'{windowID}' : unable to locate this window title");
            }
        }

        private void CheckControlVisibility(string windowTitle, string controlID, string windowText = "")
        {
            var isControlVisible = AutoItX.ControlCommand(windowTitle, windowText, controlID, "IsVisible", "");

            if (controlID == disconnectButton) //Disconnect VPN if its already connected
            {
                if (Int16.Parse(isControlVisible) == 1) //Autoit will return 1 if the control is visible
                {
                    AutoItX.ControlClick(domainWindowTitle, "", disconnectButton); //Uncomment this line if you need disconnect vpn if its already connected
                    WaitForWindow(domainWindowTitle, 10, "Ready to connect.");  //Close vpn window once it got disconnected successfully
                    AutoItX.WinClose(domainWindowTitle, "Ready to connect.");
                    Environment.Exit(0);
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
            if (checkcertificationerror == "yes") // Only wait for certification error popup when config have yes option. 
            {
                CheckForCertificationErrorPopup();
            }
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

        private void CheckForCertificationErrorPopup()
        {
            //Wait for certification error popup. This is optional because all domains wont have certificate errors
            WaitForWindow(certificationWindowTitle, 8, certificationWindowText);

            if (AutoItX.WinExists(certificationWindowTitle, "Untrusted Server Blocked!") == 1)
            {
                CheckControlVisibility(certificationWindowTitle, keepMeSafeButton, "Keep Me Safe");
                ShowErrorPopup("Fix certification issue and rerun the script..");
            }
            else if (AutoItX.WinExists(certificationWindowTitle, certificationWindowText) == 1)
            {
                CheckControlVisibility(certificationWindowTitle, connectAnywayButton);
                AutoItX.ControlClick(certificationWindowTitle, "", connectAnywayButton);
            }

        }

        private void LoginUsingUserCredentials()
        {
            WaitForWindow(loginWindowTitle);

            //Check all elements are visible in this particular window
            CheckControlVisibility(loginWindowTitle, usernameTextbox);
            CheckControlVisibility(loginWindowTitle, passwordTextbox);
            CheckControlVisibility(loginWindowTitle, okButton);

            CheckforGroupSelector(loginWindowTitle, groupDropdown);
            AutoItX.ControlSetText(loginWindowTitle, "", usernameTextbox, userName);
            AutoItX.ControlSetText(loginWindowTitle, "", passwordTextbox, passWord);
            AutoItX.ControlClick(loginWindowTitle, "", okButton);
        }

        // Check for terms and conditions popup and click accept button if any
        private void AcceptTermsandConditionPopup()
        {
            WaitForWindow(termsandConditonWindowTitle, 4); // wait for 4 secs
            if (AutoItX.WinExists(termsandConditonWindowTitle) == 1)
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
                    string defaultGroupName = AutoItX.ControlCommand(windowTitle, "", groupDropdown, "GetCurrentSelection", groupName); // get the default group name if any

                    if (defaultGroupName != groupName || string.IsNullOrEmpty(defaultGroupName))
                    {
                        AutoItX.ControlCommand(windowTitle, "", groupDropdown, "SelectString", groupName);
                        string selectedGroupName = AutoItX.ControlCommand(windowTitle, "", groupDropdown, "GetCurrentSelection", groupName);


                        if (selectedGroupName != groupName)
                            ShowErrorPopup($"Unable to find 'GROUP: {groupName}'. Check it in CONFIG file.");


                        //After selecting a value from dropdown, popup window becomes disbles.. wait till it becomes active again
                        WaitTillControlEnabled(loginWindowTitle, okButton);
                    }
                }
                else
                    ShowErrorPopup($"Enter valid GROUP {groupName} inside CONFIG file");
            }
        }
    }
}
