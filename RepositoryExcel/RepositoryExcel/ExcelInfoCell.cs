using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
namespace RepositoryExcel
{
    public class ExcelInfoCell:IExcelInfo
    {
        protected string name;
        protected string formula;
        protected string value;
        protected int size;

        public ExcelInfoCell()
        {
            this.name = "新細明體";
            this.size = 12;
            this.value = "";
        }

        public int Size { get { return this.size; } set { this.size = value; } }
        public string FontName { get { return this.name; } set { this.name = value; } }
        public string Value{ get { return this.value; } set { this.value = value; } }


        public int Y { get; set; }

        public int X { get; set; }




        /// <summary> 字呈現  
        /// 字體顏色
        /// 背景顏色
        /// 尺寸
        /// 字型
        /// 粗體
        /// 格式
        /// 對齊
        public Color fontColor { get; set; }
        public Color backColor { get; set; }
       
        
        public bool IsBold { get; set; }
        public string Formal { get; set; }
        public XlHAlign xlHAlign { get; set; }
        public string Formula { get { return this.formula; } set { this.name = formula; } }


        #region 打印時進行座標中列注入
        //public void RowFormula(int row)
        //{


        //}
        //public void SumFormula(int rowA, int rowB)
        //{

        //}
        #endregion


    }
}
