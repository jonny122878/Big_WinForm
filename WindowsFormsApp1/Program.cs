using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using System.Configuration;
using WindowsFormsApp1.Algorithm.Models;
using WindowsFormsApp1.Algorithm;

namespace WindowsFormsApp1
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static async Task Main()
        {
           
                var server = "wpdb2.hihosting.hinet.net";
                var acct = "p89880749_p89880749";
                var pwd = "Jonny1070607!@#$%";
                var dbName = "p89880749_test";
                var serverPwd = DESFunction.DESEncrypt(server, ProjectSet.PASSWORDKEY);
                var acctPwd = DESFunction.DESEncrypt(acct, ProjectSet.PASSWORDKEY);
                var pwdPwd = DESFunction.DESEncrypt(pwd, ProjectSet.PASSWORDKEY);
                var dbNamePwd = DESFunction.DESEncrypt(dbName, ProjectSet.PASSWORDKEY);
                var serverPrase = DESFunction.DESDecrypt("5CE47D09F1577500C585A9BE1535D91E21EB8251D6CC04E0392AAFBD2EA4DFB3", ProjectSet.PASSWORDKEY);
                var acctPrase = DESFunction.DESDecrypt("39C359819BA80DD8F7F11655A03403CAEBAD4DC6489BBD7C", ProjectSet.PASSWORDKEY);
                var pwdPrase = DESFunction.DESDecrypt("3466534F287AB753588B2EB865DE1962D2B773642A2885E4", ProjectSet.PASSWORDKEY);
                var dbNamePrase = DESFunction.DESDecrypt("39C359819BA80DD8EF774CE538DA11D8", ProjectSet.PASSWORDKEY);
                var input = "testB";
                var inputPwd = DESFunction.DESEncrypt(input, ProjectSet.PASSWORDKEY);
                Console.WriteLine("");
           

            if (Environment.OSVersion.Version.Major >= 6)  
                SetProcessDPIAware(); 

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
        [System.Runtime.InteropServices.DllImport("user32.dll")]  
        public static extern bool SetProcessDPIAware();      
    }

   

}
