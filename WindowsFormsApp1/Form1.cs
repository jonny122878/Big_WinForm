using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using log4net;
using WindowsFormsApp1.Algorithm;
using WindowsFormsApp1.ViewModels;
using System.Diagnostics;
namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private AuthMenu anuthMenu;
        private ValidLoginViewModel validLoginViewModel;
        private List<Tuple<string, Action<object, EventArgs>>> menuForms;
        private List<Button> btns;
        private ILog log = LogManager.GetLogger(typeof(Form1));
        private string _directory;
        public Form1()
        {
            InitializeComponent();

            this._directory = Directory.GetCurrentDirectory() + @"\SmallPrograms";





            var fileLog = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\log4net.config");            
            log4net.Config.XmlConfigurator.ConfigureAndWatch(fileLog);
            log.Info("程式啟動");
            anuthMenu = new AuthMenu();           

            menuForms = new List<Tuple<string, Action<object, System.EventArgs>>>
            { 
                //new Tuple<string, Action<object, System.EventArgs>>("A1",(s, e) => 
                //{
                //    A1 a1 = new A1();
                //    a1.Show();
                //}),
                new Tuple<string, Action<object, System.EventArgs>>("A4",(s, e) =>
                {
                    A4 a4 = new A4();
                    a4.Show();
                }),
                new Tuple<string, Action<object, System.EventArgs>>("C2",(s, e) =>
                {
                    Hello a4 = new Hello();
                    a4.Show();
                }),
                //milk
                new Tuple<string, Action<object, System.EventArgs>>("牛奶銷售分析",(s, e) =>
                {
                    Process exep = new Process();
                    exep.StartInfo.FileName = @"C:\程式\WindowsFormsApp1\WindowsFormsApp1\bin\Debug\小程式\MilkStatistics.exe";
                    exep.StartInfo.UseShellExecute = false;
                    exep.Start();
                    //MilkStatistics.Form1 a1 = new MilkStatistics.Form1();
                    //a1.Show();
                }),
            };            
        }

        private void 打分數ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var path = this._directory;
            if(!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            LoginForm loginForm = new LoginForm();
            
            loginForm.ShowDialog();

            if(!loginForm.IsLoin)
            {
                Application.Exit();
                this.Close();
                return;
            }


            validLoginViewModel = loginForm.TestAPIModel;

            //create Button
            btns = validLoginViewModel.DisplayAuthentications.Select((r, idx) =>
            {
                var step = idx * 300;
                var name = r.Value.Item2;
                Button btn = new Button();
                btn.Visible = true;
                btn.Enabled = r.Value.Item3;
                btn.Text = r.Value.Item1;
                btn.Size = new Size(200, 60);

                btn.Location = new Point(450 + step, 150);
                //var btnEvent = this.menuForms.First(c => c.Item1 == name).Item2;
                var productPath = this._directory + @"\" + r.Key;
                var productName = r.Value.Item1 + ".exe";
                var product = Path.Combine(productPath, productName);
                Action<object, System.EventArgs> btnEvent = (s,c) => 
                {

                    //var serverPrase = DESFunction.DESDecrypt("5CE47D09F1577500C585A9BE1535D91E21EB8251D6CC04E0392AAFBD2EA4DFB3", ProjectSet.PASSWORDKEY);
                    //var acctPrase = DESFunction.DESDecrypt("39C359819BA80DD8F7F11655A03403CAEBAD4DC6489BBD7C", ProjectSet.PASSWORDKEY);
                    //var pwdPrase = DESFunction.DESDecrypt("3466534F287AB753588B2EB865DE1962D2B773642A2885E4", ProjectSet.PASSWORDKEY);
                    //var dbNamePrase = DESFunction.DESDecrypt("39C359819BA80DD8EF774CE538DA11D8", ProjectSet.PASSWORDKEY);

                    Process exep = new Process();
                    //加入密碼驅動
                    exep.StartInfo.FileName = product;
                    exep.StartInfo.UseShellExecute = false;
                    exep.StartInfo.Arguments = "3466534F287AB753588B2EB865DE1962D2B773642A2885E4";
                    exep.Start();
                };
                btn.Click += new System.EventHandler(btnEvent);
                return btn;
            }).ToList();

            btns.ForEach(item => { this.Controls.Add(item); });

            //Button btnTest = new Button();
            //btnTest.Visible = true;
            //btnTest.Text = "A1";
            //btnTest.Size = new Size(30, 30);
            //btnTest.Location = new Point(200, 200);

            //Button btnTest1 = new Button();
            //btnTest1.Visible = true;
            //btnTest1.Text = "A1";
            //btnTest1.Size = new Size(30, 30);
            //btnTest1.Location = new Point(250, 200);

            //this.Controls.Add(btnTest);
            //this.Controls.Add(btnTest1);

            this.label1.Text = String.Format(@"Hi~{0}", validLoginViewModel.Name);


            this.主畫面ToolStripMenuItem_Click(null,null);
            //test
            #region old權限用不到
            //validLoginViewModel = new ValidLoginViewModel
            //{
            //    DisplayAuthentications = new List<Tuple<string, string, bool>>
            //    {
            //        new Tuple<string, string, bool>("A","A1",true),
            //        new Tuple<string, string, bool>("A","A2",false),
            //        new Tuple<string, string, bool>("A","A3",false),
            //        new Tuple<string, string, bool>("A","Hello",true),
            //        new Tuple<string, string, bool>("A","A4",false),
            //        new Tuple<string, string, bool>("B","B1",false),
            //        new Tuple<string, string, bool>("B","B2",false),
            //        new Tuple<string, string, bool>("C","C1",true),
            //    }
            //};

            //anuthMenu.LoadMenu(validLoginViewModel.DisplayAuthentications, menuForms, this.MainMenuStrip); 
            #endregion
        }

        private void 支援ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process process = new Process();      
            process.StartInfo.FileName = Path.Combine(Directory.GetCurrentDirectory(), "AnyDesk.exe");
            process.Start();
            //Form form = Assembly.Load().CreateInstance("WindowsFormsApp1." + "A1");
        }

        private void 小程式資訊ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Visible = true;
            this.label1.Visible = false;
            this.label2.Visible = false;
            btns.ForEach(btn => {
                btn.Visible = false;            
            });


            DataTable dt = new DataTable();
            var columns = new DataColumn[] 
            { 
                new DataColumn("程式名",typeof(string)),
                new DataColumn("到期日",typeof(DateTime)),
                new DataColumn("月費",typeof(int)),
                new DataColumn("使用狀態",typeof(string)),
            };
            dt.Columns.AddRange(columns);

            var rows = this.validLoginViewModel.OrderGoodes.Select(r => {

                var used = (r.Value.Item4) ? "可使用" : "不能用";
                var row = new object[] {r.Value.Item2,r.Value.Item1,r.Value.Item3,r.Value.Item4};
                return row;
            }).ToList();

            rows.ForEach(item => 
            {
                dt.Rows.Add(item);
            });

            this.dataGridView1.DataSource = dt;
        }

        private void 主畫面ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Visible = false;
            this.label1.Visible = true;
            this.label2.Visible = true;
            btns.ForEach(btn => {
                btn.Visible = true;
            });

        }

        private void LoadPrograms()
        {
            

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
