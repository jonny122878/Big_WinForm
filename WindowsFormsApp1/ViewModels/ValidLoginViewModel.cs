using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WindowsFormsApp1.ViewModels
{
    public class ValidLoginViewModel
    {
        public bool Succeeded { get; set; }

        public string Describe { get; set; }

        public string Name { get; set; }

        public List<KeyValuePair<string, Tuple<DateTime, string, int, bool>>> OrderGoodes { get; set; }

        //第一個項目為種類、二個名稱、三個為是否有權限
        //種類方式GROUP 父Item 名稱部分為子Item項目
        public List<Tuple<string,string, bool>> DisplayAuthentications { get;set;}
    }
}
