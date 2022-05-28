using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.Algorithm;
using WindowsFormsApp1.ViewModels;
namespace WindowsFormsApp1.Algorithm
{
    //未來擴充可能會因為多了次數關係而有使用中突然沒權限
    public class AuthMenu
    {
        protected List<Tuple<ToolStripMenuItem, List<ToolStripMenuItem>>> menus;

        public List<Tuple<ToolStripMenuItem, List<ToolStripMenuItem>>> Menus { get => menus; set => menus = value; }

        protected List<Tuple<ToolStripMenuItem,List<ToolStripMenuItem>>> GetMenus(List<Tuple<string, string, bool>> displayAuthentications)
        {
            return displayAuthentications.Where(r => r.Item3).
                    GroupBy(r => r.Item1).Select(r =>
                    {
                        var sort = new ToolStripMenuItem() 
                        { 
                            Name = r.Key,
                            Text = r.Key,
                            Size = new System.Drawing.Size(43, 20)
                        };
                        var items = r.Select(c =>
                        {
                            return new ToolStripMenuItem()
                            {
                                Name = c.Item2,
                                Text = c.Item2,
                                Size = new System.Drawing.Size(110, 22)
                            };
                        }).ToList();
                        return new Tuple<ToolStripMenuItem, List<ToolStripMenuItem>>(sort, items);
                    }).ToList();
        }


        public void LoadMenu(List<Tuple<string, string, bool>> displayAuthentications, List<Tuple<string, Action<object, System.EventArgs>>> menuForms, MenuStrip menuStrip)
        {

            //將外部第一層選單加入
            Menus = this.GetMenus(displayAuthentications);
            var fatherMenuArr = Menus.Select(r => r.Item1).ToArray();
            menuStrip.Items.AddRange(fatherMenuArr);

            Menus.ForEach(menu =>
            {            
                menu.Item1.DropDownItems.AddRange(menu.Item2.ToArray());
            });



            //將第二層包含Form功能表單加入
            var allMenus = menus.Select(r => r.Item2).
                SelectMany(r => r).ToList();

            //先將選單和要對應產生的form做交集綁定在一起
            var querys = from d in allMenus
                         join m in menuForms on d.Name equals m.Item1
                         select new Tuple<ToolStripMenuItem, Action<object, System.EventArgs>>(d,m.Item2);

            querys.ToList().
                ForEach(menu => 
            { 
                menu.Item1.Click += new System.EventHandler(menu.Item2);
            });



            //this.支援ToolStripMenuItem.Size = new System.Drawing.Size(43, 20); 父
            //this.簡易ToolStripMenuItem.Size = new System.Drawing.Size(110, 22); 子
        }

        public void LoadMenu(List<Tuple<string, string, bool>> displayAuthentications, MenuStrip menuStrip,ITest testDeny)
        {
            var result = testDeny.TestMethod(50);
            Menus = this.GetMenus(displayAuthentications);
            var fatherMenuArr = Menus.Select(r => r.Item1).ToArray();

            menuStrip.Items.AddRange(fatherMenuArr);
            Menus.ForEach(menu =>
            {
                var sonMenuArr = menu.Item2.ToArray();
                menu.Item1.DropDownItems.AddRange(sonMenuArr);
            });

            //this.支援ToolStripMenuItem.Size = new System.Drawing.Size(43, 20); 父
            //this.簡易ToolStripMenuItem.Size = new System.Drawing.Size(110, 22); 子
        }
    }
}
