using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using WindowsFormsApp1.ViewModels;
using WindowsFormsApp1.Algorithm;
using WindowsFormsApp1.Algorithm.Models;
namespace WindowsFormsApp1
{
    public partial class LoginForm : Form
    {
        #region API DEMO

        //using (var client = new HttpClient())
        //{
        //    var baseAddr = "https://localhost:44340/Home/TestAPI";
        //    client.BaseAddress = new Uri(baseAddr);
        //    client.DefaultRequestHeaders.Accept.Clear();
        //    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        //    // HTTP GET
        //    HttpResponseMessage response = await client.GetAsync("");
        //    //response.EnsureSuccessStatusCode();    // Throw if not a success code.
        //    if (response.IsSuccessStatusCode)
        //    {

        //        var str = await response.Content.ReadAsStringAsync();
        //        //var str = @"{'ID':'ID','Name':'NAME'}";
        //        str = str.Replace(@"\u0022", "\'");
        //        str = str.Replace("\"", "");
        //        ////JObject json = JObject.Parse(stream);
        //        TestAPIModel testAPIModel = JsonConvert.DeserializeObject<TestAPIModel>(str);
        //        Console.Write("");
        //    }
        //    //return empNames;
        //}  
        #endregion
        #region Demo MAC

        #endregion

        private bool _isLoin = false;
        private ValidLoginViewModel _testAPIModel;

        public LoginForm()
        {
            InitializeComponent();
        }

        public bool IsLoin { get => _isLoin; set => _isLoin = value; }
        public ValidLoginViewModel TestAPIModel { get => _testAPIModel; set => _testAPIModel = value; }

        private async void button1_Click(object sender, EventArgs e)
        {
            //output:loginViewModel
            #region 打包要登入驗證Json
            var netWorks = NetworkInterface.GetAllNetworkInterfaces();
            var mac = netWorks.Where(r => r.NetworkInterfaceType == NetworkInterfaceType.Ethernet).
                Select(r => r.GetPhysicalAddress()).First().ToString();

            //var randomKey = DESFunction.GetRandomStringByHashKey(8);
            var pwd = DESFunction.DESEncrypt(this.txtPwd.Text, ProjectSet.PASSWORDKEY);
            var loginViewModel = new LoginViewModel
            {
                Email = this.txtAcct.Text,
                Password = pwd,
                //PasswordKey = randomKey,
                MAC_IP = mac,
            };

            #endregion
            Console.WriteLine("");
            #region 發送功能暫時隱藏

            //using (HttpClient client = new HttpClient())
            //{
            //    client.DefaultRequestHeaders.Accept.Clear();
            //    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //    //var baseAddr = "https://localhost:44340/Home/LoginByWinForm";
            //    var baseAddr = "http://topefficiencywork.tw/Home/LoginByWinForm";
            //    var dicts = new Dictionary<string, string>();
            //    dicts.Add(nameof(loginViewModel.Email), loginViewModel.Email);
            //    dicts.Add(nameof(loginViewModel.MAC_IP), loginViewModel.MAC_IP);
            //    dicts.Add(nameof(loginViewModel.Password), loginViewModel.Password);
            //    var query = "?" + dicts.Select(r => r.Key + "=" + r.Value).
            //        Aggregate((cur, next) => cur + "&" + next);
            //    query = query.Remove(query.Length - 1, 1);
            //    HttpResponseMessage response = await client.GetAsync(baseAddr + query);
            //    if (response.IsSuccessStatusCode)
            //    {
            //        var str = await response.Content.ReadAsStringAsync();
            //        str = str.Replace(@"\u0022", "\'");
            //        str = str.Replace("\"", "");
            //        _testAPIModel = JsonConvert.DeserializeObject<ValidLoginViewModel>(str);
            //    }
            //}

            #endregion


            #region 發送API到站台
            //using (System.Net.Http.HttpClient client = new HttpClient())
            //{
            //    //var baseAddr = new Uri("https://localhost:44340/Home/LoginByWinForm");                
            //    //client.BaseAddress = baseAddr;
            //    client.DefaultRequestHeaders.Accept.Clear();
            //    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //    var baseAddr = "https://localhost:44340/Home/LoginByWinForm";

            //    var dicts = new Dictionary<string,string>();
            //    dicts.Add(nameof(loginViewModel.Email), loginViewModel.Email);
            //    dicts.Add(nameof(loginViewModel.MAC_IP), loginViewModel.MAC_IP);
            //    dicts.Add(nameof(loginViewModel.Password), loginViewModel.Password);

            //    var query = "?" + dicts.Select(r => r.Key + "=" + r.Value).
            //        Aggregate((cur, next) => cur + "&" + next);               
            //    query = query.Remove(query.Length -1, 1);
            //    HttpResponseMessage response = await client.GetAsync(baseAddr + query);
            //    // HTTP GET
            //    //HttpResponseMessage response = await client.GetAsync(baseAddr + query);
            //    ////response.EnsureSuccessStatusCode();    // Throw if not a success code.
            //    if (response.IsSuccessStatusCode)
            //    {

            //        var str = await response.Content.ReadAsStringAsync();
            //        //var str = @"{'ID':'ID','Name':'NAME'}";
            //        str = str.Replace(@"\u0022", "\'");
            //        str = str.Replace("\"", "");
            //        _testAPIModel = JsonConvert.DeserializeObject<ValidLoginViewModel>(str);
            //        Console.Write("");
            //    }
            //return empNames;
            //}
            #endregion
            //test
            this._testAPIModel = new ValidLoginViewModel
            { Succeeded = true,
                Name = "礎謙",
                OrderGoodes = new List<KeyValuePair<string, Tuple<DateTime, string, int, bool>>> 
                { 
                    new KeyValuePair<string, Tuple<DateTime, string, int, bool>>("牛奶銷售分析",new Tuple<DateTime, string, int, bool>(new DateTime(2023,1,1),"牛奶銷售分析",100,true))
                },
            };

            if (!_testAPIModel.Succeeded)
            {
                MessageBox.Show("無此帳號密碼");
            }

            if(_testAPIModel.Succeeded)
            {
                this._isLoin = true;
                this.Close();
            }


        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
