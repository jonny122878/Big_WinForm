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

        //訂單編號、名稱、版本、是否到期
        public List<KeyValuePair<string, Tuple<string, string, bool>>> DisplayAuthentications { get; set; }
    }
}
