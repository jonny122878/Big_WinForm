using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using WindowsFormsApp1.ViewModels;
using WindowsFormsApp1.Algorithm;
using WindowsFormsApp1.Algorithm.Models;
using System.Configuration;
using System.Diagnostics;

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
        private string _baseAddr;
        public LoginForm()
        {
            InitializeComponent();
        }
        private void LoginForm_Load(object sender, EventArgs e)
        {
            this._baseAddr = ConfigurationManager.AppSettings["WebSite"];
            if (ConfigurationManager.AppSettings["IsTestConnect"] == "1")
            {
                this.txtAcct.Text = @"test0906568135@gmail.com";
                this.txtPwd.Text = "0906568135";
                //this._baseAddr = "https://localhost:44340/";
                //this.button1_Click(null,null);
            }
        }
        public bool IsLoin { get => _isLoin; set => _isLoin = value; }
        public ValidLoginViewModel TestAPIModel { get => _testAPIModel; set => _testAPIModel = value; }


        private async Task<ValidLoginViewModel> ValidActPwd()
        {
            this.button1.Enabled = false;
            this.txtAcct.Enabled = false;
            this.txtPwd.Enabled = false;
            this.label4.Visible = true;
            this.label4.Text = "驗證帳號密碼中";
            //output:loginViewModel
            #region 網卡資料,密碼解析 => class Model 打包要登入驗證Json
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
            #region class Model => Json => 發送API到站台 => Get Json => Prase class Model
            using (System.Net.Http.HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //var baseAddr = "https://localhost:44340/Download/LoginByWinForm";
                var baseAddr = this._baseAddr + "Download/LoginByWinForm";
                var dicts = new Dictionary<string, string>();
                dicts.Add(nameof(loginViewModel.Email), loginViewModel.Email);
                dicts.Add(nameof(loginViewModel.MAC_IP), loginViewModel.MAC_IP);
                dicts.Add(nameof(loginViewModel.Password), loginViewModel.Password);

                var query = "?" + dicts.Select(r => r.Key + "=" + r.Value).
                    Aggregate((cur, next) => cur + "&" + next);
                //query = query.Remove(query.Length - 1, 1);
                //HttpResponseMessage response = await client.GetAsync(baseAddr + query);
                HttpResponseMessage response = await client.GetAsync(baseAddr + query);
                //HttpResponseMessage response = await client.GetAsync("https://www.topefficiencywork.tw/Download/Index");
                // HTTP GET
                //HttpResponseMessage response = await client.GetAsync(baseAddr + query);
                response.EnsureSuccessStatusCode();    // Throw if not a success code.
                ///
                if (response.IsSuccessStatusCode)
                {

                    var str = await response.Content.ReadAsStringAsync();
                    //var str = @"{'ID':'ID','Name':'NAME'}";
                    str = str.Replace(@"\u0022", "\'");
                    str = str.Replace("\"", "");
                    _testAPIModel = JsonConvert.DeserializeObject<ValidLoginViewModel>(str);
                    Console.Write("");
                }
            }
            #endregion

            return _testAPIModel;
        }

        private async Task<bool> UpdatePrograms(List<string> downloadProucts)
        {
            this.label4.Visible = false;
            this.label5.Visible = true;
            this.label5.Text = String.Format(@"小程式有{0}隻需更新中",downloadProucts.Count);
            //下載更新小程式
            using (System.Net.Http.HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //var baseAddr = "https://localhost:44340/Download/UpdateVersion";
                var baseAddr = this._baseAddr + "Download/UpdateVersion";

                UpdateVersionViewModel updateVersionViewModel = new UpdateVersionViewModel
                {
                    User = this.txtAcct.Text.Split('@')[0],
                    Versions = downloadProucts
                };



                var query = "?User=" + this.txtAcct.Text.Split('@')[0];
                var queryVersions = downloadProucts.Aggregate((cur, next) => cur + "&Versions=" + next).ToString();
                query = query + "&Versions=" + queryVersions;
                //query = query + downloadProucts.Aggregate((cur, next) => cur + "&" + next).ToString();
                HttpResponseMessage response = await client.GetAsync(baseAddr + query);

                if (response.IsSuccessStatusCode)
                {
                    var bytes = await response.Content.ReadAsByteArrayAsync();
                    var fileName = response.Content.Headers.ContentDisposition.FileName;
                    var file = Path.Combine(Directory.GetCurrentDirectory() + @"\SmallPrograms", fileName);
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                    using (var stream = new FileStream(file, FileMode.CreateNew, FileAccess.Write))
                    {
                        await stream.WriteAsync(bytes, 0, bytes.Length);
                    }
                }
            }
            var batch = "PraseZip.bat";
            Process exep = new Process();
            exep.StartInfo.FileName = Path.Combine(Directory.GetCurrentDirectory(), batch);
            exep.StartInfo.UseShellExecute = false;
            exep.StartInfo.CreateNoWindow = true;
            exep.Start();

            return true;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            //this.button1
            //this.label4
            Console.WriteLine("");

            this._testAPIModel = await this.ValidActPwd();

            //test
            #region test串接API
            if (ConfigurationManager.AppSettings["IsTest"] == "1")
            {
                //this._testAPIModel = new ValidLoginViewModel
                //{
                //    Succeeded = true,
                //    Name = "礎謙",
                //    DisplayAuthentications = new List<KeyValuePair<string, Tuple<string, bool>>>
                //    {
                //        new KeyValuePair<string, Tuple<string, bool>>("bayonet79921_0001",new Tuple<string, bool>("牛奶分析",true)),
                //        new KeyValuePair<string, Tuple<string, bool>>("bayonet79921_0002",new Tuple<string, bool>("牛奶分析A",false)),
                //        new KeyValuePair<string, Tuple<string, bool>>("bayonet79921_0003",new Tuple<string, bool>("MiniERP",true)),
                //        new KeyValuePair<string, Tuple<string, bool>>("bayonet79921_0004",new Tuple<string, bool>("MiniERP",false)),
                //    }

                //};
            } 
            #endregion
            

            if (!_testAPIModel.Succeeded)
            {
                MessageBox.Show("無此帳號密碼");
                this.txtAcct.Enabled = true;
                this.txtPwd.Enabled = true;
                this.button1.Enabled = true;
                this.label4.Visible = false;
                return;
            }


                #region 本地端小程式版本
                DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory() + "\\SmallPrograms");
                var smallDirectories = dir.GetDirectories();

                var productVersions = smallDirectories.Select(r =>
                {
                    var exe = r.GetFiles("*.exe").FirstOrDefault();
                    if (exe == null) { return new KeyValuePair<string, string>(r.Name, ""); }
                    FileVersionInfo myFileVersionInfo = FileVersionInfo.GetVersionInfo(exe.FullName);

                    return new KeyValuePair<string, string>(r.Name, myFileVersionInfo.FileVersion);
                }).ToDictionary(c => c.Key, c => c.Value);
                #endregion

                #region 遠端本地小程式版本差異比對
                List<string> downloadProucts = this._testAPIModel.DisplayAuthentications.GroupJoin(productVersions,
                            inner => inner.Key,
                            outer => outer.Key,
                            (inner, outer) =>
                            {
                                var version = outer.Select(c => c.Value).DefaultIfEmpty("").First();
                                if (version == "")
                                {
                                    return inner.Key;
                                }

                                var localVersion = Convert.ToInt32(version.Split('.').Last());
                                var webVersion = Convert.ToInt32(inner.Value.Item3.Split('.').Last());
                                var product = (webVersion > localVersion) ? inner.Key : "";
                                return product;
                            }).Where(r => r != "").ToList();
                #endregion
                Console.Write("");
                if(downloadProucts.Count == 0)
                {
                    this._isLoin = true;
                    this.Close();
                    return;
                }

                bool isUpdate = await this.UpdatePrograms(downloadProucts);
                    
                if(isUpdate)
                {
                    this._isLoin = true;
                    this.Close();
                }
        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Console.WriteLine("");
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
