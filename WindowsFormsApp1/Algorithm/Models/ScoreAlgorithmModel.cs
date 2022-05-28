using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Algorithm.Models
{
    public class ScoreAlgorithmModel
    {
        public int PeopleQty { get; set; }
        public decimal Avg_R_Top { get; set; } //取最高分群最高分

        public decimal Avg_R_H { get; set; }//取中分群最高分
        public decimal Sum { get; set; }
        public decimal HighGroupSum { get; set; }
        public decimal MiddenGroupSum { get; set; }
        public decimal LowGroupSum { get; set; }

    }
}
