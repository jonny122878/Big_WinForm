using System;
using System.Collections.Generic;
using System.Text;

namespace RepositoryExcel.ViewModels
{
    //No	Student	Midterm	Final	In_Class	Homework	Average	Average_R	Adjusted

    public class ScoreResultExcelModel
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
