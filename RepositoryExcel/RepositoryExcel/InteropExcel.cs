using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
namespace RepositoryExcel
{
    public abstract class InteropExcel: IReadAndWriteForExcel
    {
        protected List<List<Tuple<int, int, IExcelInfo> >> coordinates;
        public List<List< Tuple<int, int, IExcelInfo> >> Coordinates 
        {
            get { return this.coordinates; }
                set
                {
                    this.coordinates = value;
                } 
        }
        //不會每次都用到的背景屬性移到這裡
        //和江討論屬性和欄位的使用是否再於初始化 sheet title border
        private int sheet =1;
        private bool isBorder = true;
        private bool isTitle = true;
        
        public int ActiveSheet
        {
            get { return sheet; }
            set { sheet = value; }
        }
        public bool IsBorder
        {
            get { return this.isBorder; }
            set { this.isBorder = value; }
        }
        public bool IsTitle
        {
            get { return this.isTitle; }
            set { this.isTitle = value; }
        }
        public List<IExcelInfo> Titles { get; set; }
        

        protected Microsoft.Office.Interop.Excel.Application excelApp;
        protected Microsoft.Office.Interop.Excel._Workbook wBook;
        protected Microsoft.Office.Interop.Excel._Worksheet wSheets;
       
       

        public InteropExcel(string file, bool WhVisible = true)
        {
            //開啟工作表
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            // 讓Excel文件可見
            excelApp.Visible = WhVisible;
            // 停用警告訊息
            excelApp.DisplayAlerts = false;
            // 設定工作表焦點
            wBook = excelApp.Workbooks.Open(file);            
            wBook.Activate();
        }

        #region 打印表格邏輯
        //Model map ExcelInfo
        //將其轉印成帶有x,y座標的Enumerable
        //是否要打印title轉成Enumerable(並插入)
        //將excelInfo打印的呈現到Excel
        //進行Range的呈現
        //是否將表格的框線給列印
        #endregion

        protected List<List<IExcelInfo>> GetexcelInfies<TEntity>(List<TEntity> data,Func<TEntity, List<IExcelInfo>> map)
        {
            return data.Select(map).ToList();
        }
        protected void Loadcoodrinates(List<List<IExcelInfo>> excelInfies, int y, int x)
        {
            y = y - 1;//從此列開始打印起點，所以要倒退扣回到這位置
            x = x - 1;
            this.coordinates = excelInfies.Select((row,rowIndex) =>
            row.Select(
                (col, colIndex) =>
            new Tuple<int, int, IExcelInfo>
                (y + col.Y + rowIndex, x + col.X, col)
            ).ToList()
            ).ToList();
        }
        protected void LoadTitle<TEntity>(List<List<IExcelInfo>> excelInfies,int y,int x)
        {
            Type titleType = typeof(TEntity);
            PropertyInfo[] prArr = titleType.GetProperties();
            var title = prArr.Select((p,index)=> new Tuple<int, int, IExcelInfo>
                            (y, x + index, new ExcelInfoCell { Value = p.Name })
            ).ToList();
            this.coordinates.Insert(0, title);
        }
        
        protected List<Tuple<int, int, IExcelInfo>> GetcellEnumerable()
        {
            return this.coordinates.SelectMany(e => e).ToList();
        }
        protected void WriteInfo(List<Tuple<int, int, IExcelInfo>> cellEnumerable)
        {
            cellEnumerable.ForEach(item =>
            {
                this.wBook.ActiveSheet.Cells[item.Item1, item.Item2].Value = item.Item3.Value;
                this.wBook.ActiveSheet.Cells[item.Item1, item.Item2].Font.Size = item.Item3.Size;
                this.wBook.ActiveSheet.Cells[item.Item1, item.Item2].Font.Name = item.Item3.FontName;

            });
        }

        protected void PrintBorder(List<Tuple<int, int, IExcelInfo>> cellEnumerable)
        {
            var y1 = cellEnumerable.Min(e => e.Item1);
            var y2 = cellEnumerable.Max(e => e.Item1);
            var x1 = cellEnumerable.Min(e => e.Item2);
            var x2 = cellEnumerable.Max(e => e.Item2);
            Range range = this.wBook.Sheets[sheet].Range(wBook.Sheets[sheet].Cells[y1, x1], wBook.Sheets[sheet].Cells[y2, x2]);
            range.Merge(true);
        }

        public void WriteToTable<TEntity>(int y, int x, List<TEntity> data, Func<TEntity, List<IExcelInfo>> map) where TEntity : class, new()
        {
            #region 原始實現
            ////Model map ExcelInfo
            //var excelInfies = data.Select(map).ToList();
            ////將其轉印成帶有x,y座標的Enumerable
            //var coodrinates = excelInfies.Select(row =>
            //row.Select((col, colIndex) =>
            //new Tuple<int, int, IExcelInfo>(y + col.Y, x + col.X + colIndex, col)
            //)
            //).ToList();
            ////是否要打印title轉成Enumerable(並插入)
            //if (this.isTitle == true)
            //{
            //    Type titleType = typeof(TEntity);
            //    PropertyInfo[] prArr = titleType.GetProperties();
            //    var title = prArr.Select(p => new Tuple<int, int, IExcelInfo>
            //                    (y, x, new ExcelInfoCell { Value = p.Name })
            //    ).ToList();
            //    coodrinates.Insert(0, title);
            //}
            ////展開成cell
            //var cellEnumerable = coodrinates.SelectMany(e => e).ToList();


            ////將excelInfo打印的呈現到Excel
            //cellEnumerable.ForEach(item =>
            //{
            //    this.wBook.ActiveSheet.Cells[item.Item1, item.Item2].Value = item.Item3.Value;
            //    this.wBook.ActiveSheet.Cells[item.Item1, item.Item2].Font.Size = item.Item3.Size;
            //    this.wBook.ActiveSheet.Cells[item.Item1, item.Item2].Font.Name = item.Item3.FontName;

            //});
            ////進行Range的呈現(暫時省略)

            ////是否將表格的框線給列印
            //if (this.isBorder && this.coordinates.Any())
            //{
            //    var y1 = cellEnumerable.Min(e => e.Item1);
            //    var y2 = cellEnumerable.Max(e => e.Item1);
            //    var x1 = cellEnumerable.Min(e => e.Item2);
            //    var x2 = cellEnumerable.Max(e => e.Item2);
            //    Range range = this.wBook.Sheets[sheet].Range(wBook.Sheets[sheet].Cells[y1, x1], wBook.Sheets[sheet].Cells[y2, x2]);
            //    range.Merge(true);
            //}
            #endregion


            //Model map ExcelInfo
            var excelInfies = this.GetexcelInfies<TEntity>(data,map);
            //將其轉印成帶有x,y座標的Enumerable
            

            this.Loadcoodrinates(excelInfies, y, x);
            //是否要打印title轉成Enumerable(並插入)
            if (this.isTitle)
            {
                this.LoadTitle<TEntity>(excelInfies, y, x);
            }
            //將excelInfo打印的呈現到Excel
            var cellEnumerable = this.GetcellEnumerable();
            //進行Range的呈現(暫時省略)
            this.WriteInfo(cellEnumerable);
            //是否將表格的框線給列印
            //if(this.isBorder  && this.coordinates.Any())
            //{
            //    this.PrintBorder(cellEnumerable);
            //}
        }
        //省略
        public void WriteToRange<TEntity>(int y1, int x1, int y2, int x2, IExcelInfo data)
        {
            throw new NotImplementedException();
        }
          
        public void Delete(Tuple<int,int,int,int> coodRange)
        {
            var y1 = coodRange.Item1;
            var x1 = coodRange.Item2;
            var y2 = coodRange.Item3;
            var x2 = coodRange.Item4;
            Range range = this.wBook.Sheets[sheet].Range(
                wBook.Sheets[sheet].Cells[y1, x1],
                wBook.Sheets[sheet].Cells[y2, x2]);
            range.Delete();

        }
        public void Save(bool isClose = true) 
        {
            
            wBook.Save();
            if(isClose)
            {
                wBook.Close();
            }            
            //this.excelApp = null;
            //this.wBook = null;            
        }

        public void WriteToCell<TEntity>(IExcelInfo data)
        {
            if (data.Value != "")
            {
                this.wBook.ActiveSheet.Cells[data.Y, data.X].Value = data.Value;
            }
            else
            {
                this.wBook.ActiveSheet.Cells[data.Y, data.X].Formula = data.Formula;
            }

            this.wBook.ActiveSheet.Cells[data.Y, data.X].Font.Size = data.Size;
            this.wBook.ActiveSheet.Cells[data.Y, data.X].Font.Name = data.FontName;
            //this.wBook.ActiveSheet.Cells[item.Item1, item.Item2].Font.Size = item.Item3.Size;
            //this.wBook.ActiveSheet.Cells[item.Item1, item.Item2].Font.Name = item.Item3.FontName;
        }

        public void WriteOverride()
        {
//# 調寬度
//            ws.column_dimensions['A'].width = 3.88  #A欄的寬度改成3.88
//    ws.column_dimensions['L'].width = 17.5  #L欄的寬度改成17.5

//    column = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']  #B~I欄寬度改11
//    for i in column: 
//        ws.column_dimensions[i].width = 11
            //Columns("A:A").ColumnWidth = 18.75
            var lists = new List<Tuple<string, double>>             
            { 
                new Tuple<string, double>("A:A",3.88),
                new Tuple<string, double>("B:B",11),
                new Tuple<string, double>("C:C",11),
                new Tuple<string, double>("D:D",11),
                new Tuple<string, double>("E:E",11),
                new Tuple<string, double>("F:F",11),
                new Tuple<string, double>("G:G",11),
                new Tuple<string, double>("H:H",11),
                new Tuple<string, double>("I:I",17.5),
            };
            lists.ForEach(item => 
            {
                this.wBook.ActiveSheet.Columns[item.Item1].ColumnWidth = item.Item2;
            });
        }
    }
}
