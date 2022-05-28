using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
namespace RepositoryExcel
{
    public interface IExcelInfo
    {
        string Value { get; set; }
        /// <summary> 字呈現  
        /// 字體顏色
        /// 背景顏色
        /// 尺寸
        /// 字型
        /// 粗體
        /// 格式
        /// 對齊
        Color fontColor { get; set; }
        Color backColor { get; set; }
        int Size { get; set; }
        string FontName { get; set; }
        bool IsBold { get; set; }
        string Formal { get; set; }
        XlHAlign xlHAlign { get; set; }

        /// <summary> Range部分
        /// 公式
        /// 合併Key Merge
        /// 隱藏
        string Formula { get; set; }
      
        int Y { get; set; }

        int X { get; set; }

    }
}
