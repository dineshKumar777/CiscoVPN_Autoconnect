using System;
using AutoIt;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Linq;

namespace Autoit_CiscoVPN
{
    class Program
    {
        static void Main(string[] args)
        {
            //Remove depedency for AutoItX3_Assembly.dll
            AppDomain.CurrentDomain.AssemblyResolve += (sender, arg) => { if (arg.Name.StartsWith("AutoItX3")) return Assembly.Load(Properties.Resources.AutoItX3_Assembly); return null; };

            Console.WriteLine("-- Auto login script started");
            AutoLogin app = new AutoLogin();
            Console.WriteLine("-- Auto login script completed successfully");
        }
    }
}
