using AutoIt;
using System;
using System.IO;
using System.Windows.Forms;

namespace Autoit_CiscoVPN
{
    class AutoLogin : AppConfig
    {
        int waitTime = 10;
        string messageBoxTitle = "CiscoVPN_Autoit MSG";

        #region Cisco_MainWindow handles
        string domainWindowTitle = "Cisco AnyConnect Secure Mobility Client";
        string domainTextbox = "[CLASS:Edit; INSTANCE:1]";
        string domainDropdown = "[CLASS:ComboBox; INSTANCE:1]";
        string connectButton = "[CLASS:Button; TEXT:Connect; INSTANCE:1]";
        string disconnectButton = "[CLASS:Button; TEXT:Disconnect; INSTANCE:1]";
        #endregion

        #region Cisco_loginWindow handles
        string LoginWindowTitle = "Cisco AnyConnect | " + domainName;
        string groupDropdown = "[CLASS:ComboBox; INSTANCE:1]";
        string usernameTextbox = "[CLASS:Edit; INSTANCE:1]";
        string passwordTextbox = "[CLASS:Edit; INSTANCE:2]";
        string okButton = "[CLASS:Button; TEXT:OK; INSTANCE:1]";
        #endregion

        #region Common methods
        private void ShowErrorPopup(string errorMsg)
        {
            errorMsg += "\n\nScript will STOP now. Verify the error and restart the script.";
            DialogResult dialog = MessageBox.Show(errorMsg, messageBoxTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (dialog == DialogResult.OK)  //Close console window after clicking ok button
                Environment.Exit(0);
        }

        private void WaitForWindow(string windowID)
        {
            var handle = AutoItX.WinWaitActive(windowID, "", waitTime);

            if (AutoItX.WinExists(windowID) == 0) //Autoit return 1 for success and 0 for failure
                ShowErrorPopup($"'{windowID}' : unable to locate this window title");
        }

        private void CheckControlVisibility(string windowTitle, string controlID)
        {
            var isControlVisible = AutoItX.ControlCommand(windowTitle, "", controlID, "IsVisible", "");

            if(controlID == disconnectButton)
            {
                if (Int16.Parse(isControlVisible) == 1) //Autoit will return 1 if the control is visible
                    ShowErrorPopup($"Cisco VPN is ALREADY CONNECTED.");
            }
            else
            {
                if (Int16.Parse(isControlVisible) == 0) //Autoit will return 1 if the control is visible
                    ShowErrorPopup($"CONTROL NOT VISIBLE\nwindowtitle={windowTitle}\ncontrolID={controlID}");
            }
        }
        #endregion


        public AutoLogin()
        {
            OpenCiscoVPN();
            ConnectToDomain();
            LoginUsingUserCredentials();
        }
        
       

        private void OpenCiscoVPN()
        {
            if (File.Exists(ciscoExePath))
                AutoItX.Run(ciscoExePath, "", 1);
            else
                ShowErrorPopup("'Install CiscoVPN' or 'Enter the correct path in CONFIG file'.");
        }


        private void ConnectToDomain()
        {
            WaitForWindow(domainWindowTitle);

            //Check all elements are visible in particular window
            CheckControlVisibility(domainWindowTitle, domainTextbox);
            CheckControlVisibility(domainWindowTitle, domainDropdown);
            CheckControlVisibility(domainWindowTitle, disconnectButton);

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
            WaitForWindow(LoginWindowTitle);

            //Check all elements are visible in this particular window
            CheckControlVisibility(LoginWindowTitle, usernameTextbox);
            CheckControlVisibility(LoginWindowTitle, passwordTextbox);
            CheckControlVisibility(LoginWindowTitle, okButton);

            CheckforGroupSelector(LoginWindowTitle, groupDropdown);

            AutoItX.ControlSetText(LoginWindowTitle, "", usernameTextbox, userName);
            AutoItX.ControlSetText(LoginWindowTitle, "", passwordTextbox, passWord);
            AutoItX.ControlClick(LoginWindowTitle, "", okButton);
            Console.WriteLine("-- Finished login");

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
