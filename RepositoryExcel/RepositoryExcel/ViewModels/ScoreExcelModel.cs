using System;
using System.Collections.Generic;
using System.Text;

namespace RepositoryExcel.ViewModels
{
    //No.	學生	 期中測驗 	 期末測驗 	 上課表現 	 作業成績 	 原始分數加權平均 	 原始分數(四捨五入) 	 手動調整 

    public class ScoreExcelModel
    {
        public string No { get; set; }
        public string Student { get; set; }
        public string Midterm { get; set; }
        public string Final { get; set; }
        public string In_Class { get; set; }
        public string Homework { get; set; }
        public string Average { get; set; }
        public string Average_R { get; set; }
        public string Adjusted { get; set; }
    }
}
