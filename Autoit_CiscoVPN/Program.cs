using System;
using AutoIt;
using System.Configuration;
using System.IO;
using System.Windows.Forms;

namespace Autoit_CiscoVPN
{
    class Program
    {
        static void Main(string[] args)
        {
            AutoLogin app = new AutoLogin();

            Console.WriteLine("-- Auto login script started");
            app.OpenCiscoVPN();
            app.ConnectToDomain();
            app.LoginUsingUserCredentials();
            Console.WriteLine("-- Auto login script completed successfully");
        }
    }
}
