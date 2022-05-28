using System;
using System.Collections.Generic;
using System.Text;

namespace WindowsFormsApp1.ViewModels
{
    public class ScoreViewModel
    {
        public int HighGroup {get;set;}
        public int LowGroup { get; set; }

        public decimal FirstOffset { get; set; }

        public decimal AvgOffset { get; set; }
        public string Excel { get; set; }
    }
}
