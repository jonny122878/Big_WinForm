//using System;
//using System.Windows.Forms;
//using System.IO;
//using Microsoft.Office.Interop.Excel;
//using System.Threading;
//using System.Globalization;
//using LinqToExcel;
//using LinqToExcel.Query;
////using OfficeOpenXml;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace PictureWindowsForm.IOForFile
//{
//    public class Storage
//    {

//        public List<string> Content = new List<string>();

//        public int SheetIndex;
//        public int ReadCellX;
//        public int WriteCellX;


//        //之後要替換
//        public int ReadWriteCellX;
//        //web output excel
//        public void set_CellX(int ReadWriteCellX)
//        {
//            this.ReadWriteCellX = ReadWriteCellX;
//        }
//        public Storage(int ReadCellX = 0,
//            int WriteCellX = 0,
//            int SheetIndex = 1)
//        {
//            //Enum.GetName(typeof(tinySupplementCellX), ReadCellX);
//            this.ReadCellX = (int)ReadCellX;
//            this.WriteCellX = (int)WriteCellX;
//            this.SheetIndex = (int)SheetIndex;
//        }
//        public void set_sheet_index(int SheetIndex)
//        {
//            this.SheetIndex = SheetIndex;
//        }
//        public void set_ReadCellX(int ReadCellX)
//        {
//            this.ReadCellX = ReadCellX;
//        }
//        public void set_WriteCellX(int WriteCellX)
//        {
//            this.WriteCellX = WriteCellX;
//        }
//        //read excel
//        /*
//        public Storage(int ReadWriteCellX, int SheetIndex = 1)
//        {
//            this.ReadWriteCellX = ReadWriteCellX;
//            this.SheetIndex = SheetIndex;
//        }
//        */

//        //public void convert_data(convertData convertData)
//        //{
//        //    for (int i = 0; i < Content.Count; i++)
//        //    {
//        //        Content[i] = convertData.Invoke(Content[i]);
//        //        //MessageBox.Show(Content[i]);
//        //    }

//        //}
//        public void show_every_element()
//        {
//            for (int i = 0; i < this.Content.Count; i++)
//            {
//                MessageBox.Show(this.Content[i]);
//            }
//        }
//        public void show_count()
//        {
//            MessageBox.Show(this.Content.Count.ToString());
//        }

//    }

//    //模板一個excel基本行為

//    //讀取excel到storage
//    //讀取excel到storage陣列
//    //storage寫入excel
//    //storage寫入excel陣列

//    //私有行為(小函數化)
//    //storage指定索引填到指定行
//    //excel指定行填到storage陣列


//    //共用的部分提到父類
//    #region
//    //public class kevinExcel
//    //{
//    //    public List<List<Row>> ReadRidFloorFile(string filePath, string fileName)
//    //    {
//    //        using (var excelFile = new ExcelQueryFactory(this._CheckFile(filePath, fileName)))
//    //        {
//    //            var result = new List<List<Row>>();
//    //            var count = excelFile.GetWorksheetNames().Count();
//    //            for (int i = 0; i < count; i++)
//    //            {
//    //                var BeginData = new List<Row>();
//    //                BeginData = excelFile.Worksheet(i).ToList();
//    //                var GroundData = new List<Row>();
//    //                foreach (var Data in BeginData)
//    //                {
//    //                    var NewRow = new Row();
//    //                    string TmpStr = Data[0].Value.ToString();
//    //                    int CutPtn = TmpStr.IndexOf("號");
//    //                    if (CutPtn != -1)
//    //                    {
//    //                        TmpStr = TmpStr.Substring(0, CutPtn+1);
//    //                    }                        
//    //                    var NewCell = new Cell(TmpStr);
//    //                    NewRow.Add(NewCell);
//    //                    bool WhEqual = false;
//    //                    for (int iEqual=0; iEqual < GroundData.Count; iEqual++)
//    //                    {
//    //                        if (GroundData.ElementAt(iEqual)[0].Value.ToString() == NewCell.Value.ToString())
//    //                        {
//    //                            WhEqual = true;
//    //                            break;
//    //                        }
//    //                    }
//    //                    if(WhEqual == false)
//    //                    {
//    //                        GroundData.Add(NewRow);
//    //                    }
//    //                }
//    //                result.Add(GroundData);
//    //            }
//    //            return result;
//    //        }
//    //    }





//    //    public List<List<Row>> ReadGroundFile(string filePath, string fileName)
//    //    {
//    //        using (var excelFile = new ExcelQueryFactory(this._CheckFile(filePath, fileName)))
//    //        {
//    //            var result = new List<List<Row>>();
//    //            var count = excelFile.GetWorksheetNames().Count();
//    //            for (int i = 0; i < count; i++)
//    //            {
//    //                var BeginData = new List<Row>();
//    //                BeginData = excelFile.Worksheet(i).ToList();
//    //                var GroundData = new List<Row>();    
//    //                foreach (var Data in BeginData)
//    //                {
//    //                    var NewRow = new Row();
//    //                    var NewCell = new Cell(Data[0].Value + "地下");
//    //                    NewRow.Add(NewCell);
//    //                    GroundData.Add(NewRow);
//    //                }
//    //                result.Add(GroundData);
//    //            }                    
//    //            return result;
//    //        }            
//    //    }


//    //    public List<List<Row>> ReadFile(string filePath, string fileName)
//    //    {
//    //        using (var excelFile = new ExcelQueryFactory(this._CheckFile(filePath, fileName)))
//    //        {
//    //            var result = new List<List<Row>>();
//    //            var count = excelFile.GetWorksheetNames().Count();
//    //            for (int i = 0; i < count; i++)
//    //                result.Add(excelFile.Worksheet(i).ToList());
//    //            return result;
//    //        }
//    //    }

//    //    public List<string> GetSheetNames(string filePath, string fileName)
//    //    {
//    //        using (var excelFile = new ExcelQueryFactory(this._CheckFile(filePath, fileName)))
//    //        {
//    //            return excelFile.GetWorksheetNames().ToList();
//    //        }
//    //    }


//    //    public void WorkBookSetToFile(
//    //        string path, string fileName, int datasInSheet, int sheetsInExcel, List<List<Row>> datas, List<string> sheetsNmae)
//    //    {
//    //        List<List<Row>> newExcel = new List<List<Row>>();
//    //        List<string> newSheetNames = new List<string>();
//    //        var sheetsIndex = 0;

//    //        foreach (var sheet in datas)
//    //        {
//    //            var check = Math.Ceiling((double)sheet.Count() / datasInSheet);
//    //            if (check == 1)
//    //            {
//    //                newSheetNames.Add(sheetsNmae[sheetsIndex]);
//    //                newExcel.Add(sheet);
//    //            }
//    //            else
//    //            {
//    //                var newSheetName = sheetsNmae[sheetsIndex];
//    //                for (int i = 1; i <= check; i++)
//    //                {
//    //                    newSheetNames.Add(newSheetName + i.ToString());
//    //                    newExcel.Add(this._SkipAndTake<Row>(i, datasInSheet, sheet));
//    //                }
//    //            }
//    //            sheetsIndex++;
//    //        }
//    //        this.SpecilWriteFile(path, fileName, sheetsInExcel, newExcel, newSheetNames);
//    //    }

//    //    private void WriteFile(string path, string fileName, List<List<Row>> datas, List<string> sheetsNmae)
//    //    {
//    //        using (ExcelPackage p = new ExcelPackage())
//    //        {
//    //            var sheetNanmeIdx = 0;
//    //            datas.ForEach(sheetData => {
//    //                ExcelWorksheet sheet = p.Workbook.Worksheets.Add(sheetsNmae[sheetNanmeIdx]);
//    //                //var colIdx = sheetData.First().ColumnNames.Count();
//    //                var colIdx = 1;
//    //                var rowIdx = 2;
//    //                foreach (var rowData in sheetData)
//    //                {
//    //                    for (int i = 1; i <= colIdx; i++)
//    //                        sheet.Cells[rowIdx, i].Value = rowData[colIdx - 1].Value;
//    //                    rowIdx++;
//    //                }
//    //                var a = 0;
//    //                sheetNanmeIdx++;
//    //            });
//    //            p.SaveAs(new FileInfo(this._CheckPath(path, fileName)));
//    //        }
//    //    }


//    //    private void SpecilWriteFile(string path, string fileName, int sheetsInExcel, List<List<Row>> datas, List<string> sheetsNmae)
//    //    {
//    //        var check = Math.Ceiling((double)datas.Count() / sheetsInExcel);
//    //        for (int i = 1; i <= check; i++)
//    //        {
//    //            var newFileName = fileName + i.ToString() + ".xlsx";
//    //            //new Program().WriteFile(@"D:\kevinfile\2018\C#_Console開發測試區\測試資料\Test", newFileName,
//    //            //    datas.Skip((i - 1) * sheetsInExcel).Take(sheetsInExcel).ToList(),
//    //            //    sheetsNmae.Skip((i - 1) * sheetsInExcel).Take(sheetsInExcel).ToList()
//    //            //    );
//    //                this.WriteFile(path, newFileName,
//    //                this._SkipAndTake<List<Row>>(i, sheetsInExcel, datas),
//    //                this._SkipAndTake<string>(i, sheetsInExcel, sheetsNmae)
//    //                );
//    //        }
//    //    }



//    //    private List<T> _SkipAndTake<T>(int pageCnt, int pageRow, IEnumerable<T> datas)
//    //    {
//    //        return datas.Skip((pageCnt - 1) * pageRow).Take(pageRow).ToList();
//    //    }


//    //    private string _CheckFile(string filePath, string fileName)
//    //    {
//    //        var file = Path.Combine(filePath, fileName);
//    //        if (!File.Exists(file.Trim())) // 確認有無檔案
//    //            throw new Exception("無此檔案: " + file);
//    //        return file;
//    //    }

//    //    private string _CheckPath(string path, string fileName)
//    //    {
//    //        if (!Directory.Exists(path.Trim()))
//    //            throw new Exception("無此路徑" + path);
//    //        return Path.Combine(path, fileName);
//    //    }
//    //}

//    #endregion


//    #region
//    //class openDialogExcel
//    //{
//    //    private TextBox textbox;
//    //    private Button button;
//    //    public openDialogExcel(TextBox textbox, Button button)
//    //    {
//    //        this.textbox = textbox;
//    //        this.button = button;
//    //        this.button.Click += new EventHandler(button2_Click);
//    //    }
        
//    //    private void button2_Click(object sender, EventArgs e)
//    //    {
//    //        OpenFileDialog openFileDialog1 = new OpenFileDialog();
//    //        openFileDialog1.InitialDirectory = "c:\\";
//    //        openFileDialog1.RestoreDirectory = true;
//    //        if (openFileDialog1.ShowDialog() == DialogResult.OK)
//    //        {
//    //            this.textbox.Text = "";
//    //            this.textbox.Text = openFileDialog1.FileName;
//    //        }
//    //    }
//    //}

//    #endregion


//    public class tinyExcel: InputExcel
//    {        
//        public tinyExcel(string Path_FileName, bool WhVisible = true)
//        {
//            //開啟工作表
//            excelApp = new Microsoft.Office.Interop.Excel.Application();
//            // 讓Excel文件可見
//            excelApp.Visible = WhVisible;
//            // 停用警告訊息
//            excelApp.DisplayAlerts = false;
//            // 設定工作表焦點

//            wBook = excelApp.Workbooks.Open(Path_FileName);            
                            
//        }
//        public void read_data(Storage Storage1, int EndCellX)
//        {
//            wBook.Sheets[Storage1.SheetIndex].Activate();
//            int iEnd = 2;
//            while (wBook.Sheets[Storage1.SheetIndex].Cells[iEnd, EndCellX].Value != null)
//            {
                
//                iEnd = iEnd + 1;
//            }

//            for (int j = 2; j < iEnd; j++)
//            {
//               if (Storage1.ReadWriteCellX != 0)
//               {
//                    Storage1.Content.Add(Convert.ToString(wBook.Sheets[Storage1.SheetIndex].Cells[j, Storage1.ReadCellX].Value));
//               }
//            }
//        }
//        public void write_data(Storage Storage1)
//        {
//            wBook.Sheets[Storage1.SheetIndex].Activate();
//             int j = 2;
//                //設定若讀取為0則不讀取
//                for (int k = 0; k < Storage1.Content.Count; k++)
//                {
//                    if (Storage1.WriteCellX != 0)
//                    {
//                        if (Storage1.Content[k] == null)
//                        {
//                            wBook.Sheets[Storage1.SheetIndex].Cells[j, Storage1.WriteCellX].Value = "";
//                        }
//                        else
//                        {
//                            wBook.Sheets[Storage1.SheetIndex].Cells[j, Storage1.WriteCellX].Value = Storage1.Content[k];
//                        }
//                    }
//                    j = j + 1;
//                }
//        }
//    }

//    public class arrangeRoadExcel: InputExcel
//    {

//        public arrangeRoadExcel(string Path_FileName, bool WhCreate = false, bool WhVisible = true)
//        {
//            //開啟工作表
//            excelApp = new Microsoft.Office.Interop.Excel.Application();
//            // 讓Excel文件可見
//            excelApp.Visible = WhVisible;
//            // 停用警告訊息
//            excelApp.DisplayAlerts = false;
//            // 設定工作表焦點


//            if (WhCreate == true)
//            {
//                excelApp.Workbooks.Add(Type.Missing);

//                wBook = excelApp.Workbooks[1];
//                wBook.Sheets[1].Activate();
//                wBook.Sheets[3].Delete();
//                wBook.Sheets[2].Delete();
//                try
//                {
//                    System.IO.Directory.Delete(Path_FileName);
//                }
//                catch (System.IO.IOException e)
//                {

//                }
//                wBook.SaveAs(Path_FileName);
//            }
//            else
//            {
//                wBook = excelApp.Workbooks.Open(Path_FileName);
//                wBook.Activate();
//                wBook.Sheets[1].Activate();

//                //MessageBox.Show(Convert.ToString(wBook.Sheets[1].Cells[1, 1].Font.Color));

//            }
//        }
        
//        public void write_title()
//        {


//            wBook.Sheets[1].Cells[1, 1] = "次序";
//            wBook.Sheets[1].Cells[1, 2] = "房屋品牌";
//            wBook.Sheets[1].Cells[1, 3] = "案名";
//            wBook.Sheets[1].Cells[1, 4] = "編號";
//            wBook.Sheets[1].Cells[1, 5] = "屋況";

//            wBook.Sheets[1].Cells[1, 6] = "姓名";
//            wBook.Sheets[1].Cells[1, 7] = "地址";
//            wBook.Sheets[1].Cells[1, 8] = "總價";
//            wBook.Sheets[1].Cells[1, 9] = "格局";
//            wBook.Sheets[1].Cells[1, 10] = "坪數";

//            wBook.Sheets[1].Cells[1, 11] = "屋齡";
//            wBook.Sheets[1].Cells[1, 12] = "樓別(次/數)";
//            wBook.Sheets[1].Cells[1, 13] = "種類";
//            wBook.Sheets[1].Cells[1, 14] = "外觀";
//            wBook.Sheets[1].Cells[1, 15] = "車位";


//            wBook.Sheets[1].Cells[1, 16] = "登記日期";
//            wBook.Sheets[1].Cells[1, 17] = "設定";
//            wBook.Sheets[1].Cells[1, 18] = "無標題一";
//            wBook.Sheets[1].Cells[1, 19] = "無標題二";
//            wBook.Sheets[1].Cells[1, 20] = "姓名電話";

//            wBook.Sheets[1].Cells[1, 21] = "姓名地址電話";
//            wBook.Sheets[1].Cells[1, 22] = "地址電話";
//            wBook.Sheets[1].Cells[1, 23] = "姓名戶籍地電話";
//            wBook.Sheets[1].Cells[1, 24] = "戶籍地電話";
//            wBook.Sheets[1].Cells[1, 25] = "備註";

//            wBook.Sheets[1].Cells[1, 26] = "郵遞區號";

//            wBook.Sheets[1].Cells[1, 27] = "公里數";
//            //		

//        }
//        protected void count_ptn(string Str, Storage StorageData, ref int FirstPtn, ref int EndPtn)
//        {
//            const int FIRST_ROW = 2;

//            for (int jRange = 0; jRange < StorageData.Content.Count; jRange++)
//            {
//                if (Convert.ToInt16(Str) < Convert.ToInt16(StorageData.Content[jRange]))
//                {
//                    if (jRange != 0)
//                    {
//                        FirstPtn = Convert.ToInt16(StorageData.Content[jRange - 1]);
//                        EndPtn = Convert.ToInt16(StorageData.Content[jRange]) - 1;
//                    }
//                    else
//                    {
//                        FirstPtn = FIRST_ROW;
//                        EndPtn = Convert.ToInt16(StorageData.Content[jRange]) - 1;
//                    }
//                    return;
//                }
//            }

//        }

//        //標色位置對應表格索引地址
//        //抓取結果索引用開頭位置去比對前後算出要填到的行數
//        //填入表格並且記錄目前新一行的位置
//        public void write_data(Storage[] Storage1, Storage StorageColor, Storage StorageData, Storage StorageMinute)
//        {
//            int PrintPtn = 2;
//            int FirstPtn = 0;
//            int EndPtn = 0;
//            for (int iColor = 0; iColor < StorageColor.Content.Count; iColor++)
//            {
//                this.count_ptn(StorageColor.Content[iColor], StorageData, ref FirstPtn, ref EndPtn);
//                for (int jStuff = FirstPtn; jStuff <= EndPtn; jStuff++)
//                {
//                    this.storage_stuff_excel(Storage1, jStuff, PrintPtn);
//                    if (jStuff == Convert.ToInt16(StorageColor.Content[iColor]))
//                    {
//                        const int TARGET_COLOR = 15773696;
//                        wBook.Sheets[1].Cells[PrintPtn, 7].Font.Color = TARGET_COLOR;
//                    }
//                    if (jStuff == FirstPtn)
//                    {
//                        wBook.Sheets[StorageMinute.SheetIndex].Cells[PrintPtn, StorageMinute.WriteCellX].Value = StorageMinute.Content[iColor];
//                    }

//                    PrintPtn = PrintPtn + 1;
//                }

//            }
//        }
//        public void read_data(Storage[] Storage1, int EndCellX)
//        {
//            int iEnd = 2;
//            while (wBook.Sheets[1].Cells[iEnd, EndCellX].Value != null)
//            {
//                iEnd = iEnd + 1;
//            }
//            for (int i = 0; i < Storage1.Length; i++)
//            {
//                for (int j = 2; j < iEnd; j++)
//                {
//                    if (Storage1[i].ReadCellX != 0)
//                    {
//                        Storage1[i].Content.Add(Convert.ToString(wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].ReadCellX].Value));

//                    }
//                }
//            }
//        }
//        public void read_mark_color_ptn(Storage Storage1, int TargetCellX)
//        {
//            const int TARGET_COLOR = 15773696;
//            int iCellY = 2;
//            while (wBook.Sheets[1].Cells[iCellY, TargetCellX].Value != null)
//            {
//                if (Convert.ToInt32(wBook.Sheets[1].Cells[iCellY, TargetCellX].Font.Color) == TARGET_COLOR)
//                {
//                    Storage1.Content.Add(iCellY.ToString());
//                }
//                iCellY = iCellY + 1;
//            }
//        }
        


//    }
    
//    public class OfficeExcel_
//    {
//        protected Microsoft.Office.Interop.Excel.Application excelApp;
//        protected Microsoft.Office.Interop.Excel._Workbook wBook;
//        protected Microsoft.Office.Interop.Excel._Worksheet wSheets;

//        public OfficeExcel_(string file)
//        {
//            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
//            //開啟工作表
//            excelApp = new Microsoft.Office.Interop.Excel.Application();
//            // 讓Excel文件可見
//            excelApp.Visible = true;
//            // 停用警告訊息
//            excelApp.DisplayAlerts = false;

//            wBook = excelApp.Workbooks.Open(file);
//            wBook.Activate();
            

//        }
//        public void Test(int sheet)
//        {
            
//            Microsoft.Office.Interop.Excel.Range range = wBook.Sheets[sheet].Range(wBook.Sheets[sheet].Cells[1, 1], wBook.Sheets[sheet].Cells[1, 2]);
//            range.Merge(true);
//            //wSheets.Range[wBook.Sheets[sheet].Cells[1, 1], wBook.Sheets[sheet].Cells[4, 1]].Merge();


//        }
//        public void Delete(int sheet, Tuple<int, int, int, int> coodRange)
//        {
//            wBook.Sheets[sheet].Activate();
//            var y1 = coodRange.Item1;
//            var x1 = coodRange.Item2;
//            var y2 = coodRange.Item3;
//            var x2 = coodRange.Item4;
//            Range range = this.wBook.Sheets[sheet].Range(
//                wBook.Sheets[sheet].Cells[y1, x1],
//                wBook.Sheets[sheet].Cells[y2, x2]);
//            range.Delete();
//        }


//        public void Update(int sheet,List<ColumnExcel> columnExcels)
//        {
//            wBook.Sheets[sheet].Activate();
//            columnExcels.ForEach(
//                item=>
//                {
//                    wBook.Sheets[sheet].Cells[item.Row, item.Column].Value = item.Value;
//                }
//            );             
//        }
//        public void SaveChanges()
//        {
//            wBook.Save();
//        }

//        public void Close()
//        {
//            wBook.Close();
//            this.wBook = null;
//            this.excelApp = null;
//        }



//    }
    
    
//    public class InputExcel
//    {
        
//        //建構如果為創造則產生一個sheet名字為預設
//        //建構如果不為創造則打開.並且第一個sheet激活

//        //二個storage私有方法
//        //公有方法.讀取.寫入.填色.寫標題
        
//        protected Microsoft.Office.Interop.Excel.Application excelApp;
//        protected Microsoft.Office.Interop.Excel._Workbook wBook;
//        protected Microsoft.Office.Interop.Excel._Worksheet wSheets;

//        public InputExcel(string Path_FileName = "", bool WhCreate = false, bool WhVisible = false,string SheetName = "預設")
//        {
//            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
//            //開啟工作表
//            excelApp = new Microsoft.Office.Interop.Excel.Application();
//            // 讓Excel文件可見
//            excelApp.Visible = WhVisible;
//            // 停用警告訊息
//            excelApp.DisplayAlerts = false;
//            // 設定工作表焦點
            

//            if (WhCreate == true)
//            {
//                excelApp.Workbooks.Add(Type.Missing);

//                wBook = excelApp.Workbooks[1];
//                wBook.Sheets[1].Activate();
//                int SheetCount = wBook.Sheets.Count;
//                if (SheetCount == 3)
//                {
//                    wBook.Sheets[3].Delete();
//                    wBook.Sheets[2].Delete();
//                }

//                wBook.Sheets[1].Name = SheetName;
//                try
//                {
//                    System.IO.Directory.Delete(Path_FileName);
//                }
//                catch (System.IO.IOException e)
//                {

//                }
//                wBook.SaveAs(Path_FileName);
//            }
//            else
//            {
//                wBook = excelApp.Workbooks.Open(Path_FileName);
//                wBook.Activate();
//                wBook.Sheets[1].Activate();

//                //MessageBox.Show(Convert.ToString(wBook.Sheets[1].Cells[1, 1].Font.Color));

//            }
//        }
//        public void save_excel()
//        {
//            wBook.Save();
//        }
//        public void set_visible(bool WhVisible)
//        {
//            excelApp.Visible = WhVisible;
//        }
        
//        protected void excel_stuff_storage(Storage[] Storage1,int StuffCellY)
//        {
//            for (int jStuff = 0; jStuff < Storage1.Length; jStuff++)
//            {
//                if (Storage1[jStuff].ReadWriteCellX != 0)
//                {
//                    Storage1[jStuff].Content.Add(Convert.ToString(wBook.Sheets[Storage1[jStuff].SheetIndex].Cells[StuffCellY, Storage1[jStuff].ReadWriteCellX].Value));
//                }  
//            }
//        }
//        protected void storage_stuff_excel(Storage[] Storage1, int Index, int StuffCellY)
//        {
//            const int OFFSET_ROW = 2;
//            for (int jStuff = 0; jStuff < Storage1.Length; jStuff++)
//            {
//                if (Storage1[jStuff].ReadWriteCellX != 0)
//                {
//                    wBook.Sheets[Storage1[jStuff].SheetIndex].Cells[StuffCellY, Storage1[jStuff].ReadWriteCellX].Value = Storage1[jStuff].Content[Index - OFFSET_ROW];

//                }
//            }
//        }
//        public string rtn_sheet_name(int Index)
//        {
//            return wBook.Sheets[Index].Name;
//        }
//        public int rtn_sheet_count()
//        {
//            return wBook.Sheets.Count;
//        }
//        public int rtn_data_count(int SheetIndex,int CellX)
//        {
//            int i = 1;
//            bool WhRun = true;
//            while (WhRun)
//            {
//                i++;
//                if (wBook.Sheets[SheetIndex].Cells[i, CellX].Value == null)
//                {
//                    WhRun = false;
//                }                
                
//            }
//            return i;
//        }


//        public void read_data(Storage[] Storage1,int EndPtn)
//        {
//            for (int i = 0; i < Storage1.Length; i++)
//            {
//                int j = 2;
//                //設定若讀取為0則不讀取
//                //if (Storage1[i].ReadWriteCellX != 0)
//                //{
//                    while (j <= EndPtn)
//                    {
//                        if(wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].ReadCellX].Value == null)
//                        {
//                            Storage1[i].Content.Add("");
//                        }
//                        else
//                        {
//                            Storage1[i].Content.Add(Convert.ToString(wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].ReadCellX].Value));
//                        }

                        
//                        j = j + 1;
//                    }
//                //}

//            }
//        }
//        public void read_data(Storage[] Storage1)
//        {
//            for (int i =0; i< Storage1.Length;i++)
//            {                
//                int j = 2;
//                //設定若讀取為0則不讀取
//                if (Storage1[i].ReadCellX != 0) 
//                {              
//                    while (wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].ReadCellX].Value != null)
//                    {
                        
//                        Storage1[i].Content.Add(Convert.ToString(wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].ReadCellX].Value));
//                        j = j + 1;
//                    }
//                }

//            }
//        }
//        public void write_data(Storage[] Storage1,int CellColor,int CellXCount)
//        {
//            for (int i = 0; i < Storage1.Length; i++)
//            {
//                int j = rtn_data_count(Storage1[0].SheetIndex,1);
//                //設定若讀取為0則不讀取
//                if (Storage1[i].WriteCellX != 0)
//                {
//                    for (int k = 0; k < Storage1[i].Content.Count; k++)
//                    {
//                        if (Storage1[i].Content[k] == null)
//                        {
//                            wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].WriteCellX].Value = "";
//                        }
//                        else if (Storage1[i].Content[k] == "不填入")
//                        {
//                            //不作動
//                        }
//                        else
//                        {          
//                            wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].WriteCellX].Value = Storage1[i].Content[k].ToString();
//                            stuff_interior_color(CellColor, j, CellXCount,Storage1[i].SheetIndex);
//                        }

//                        j = j + 1;
//                    }

//                }

//            }
//        }
//        public void write_data(Storage[] Storage1)
//        {
//            int j = 2;
//            for (int i = 0; i < Storage1.Length; i++)
//            {
                
                
//                //設定若讀取為0則不讀取
//                if (Storage1[i].WriteCellX != 0)
//                {
//                    j = 2;
//                    for (int k = 0; k < Storage1[i].Content.Count; k++)
//                    {
//                        if (Storage1[i].Content[k] == null)
//                        {
//                            wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].WriteCellX].Value = "";
//                        }
//                        else if (Storage1[i].Content[k] == "不填入")
//                        {
//                            //不作動
//                        }
//                        else
//                        {
//                            if (Storage1[i].Content[k].ToString().IndexOf("變色") != -1)
//                            {
//                                Storage1[i].Content[k] = Storage1[i].Content[k].ToString().Replace("變色", "");
//                                stuff_cell_interior_color(7,j, Storage1[i].WriteCellX, Storage1[i].SheetIndex);
//                                wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].WriteCellX].Value = Storage1[i].Content[k].ToString();
//                            }
//                            else
//                            {
//                                //try
//                                //{
//                                    wBook.Sheets[Storage1[i].SheetIndex].Cells[j, Storage1[i].WriteCellX].Value = Storage1[i].Content[k].ToString();
//                                //}
//                                //catch
//                                //{

//                                //}                                
//                            }
                            
//                        } 

//                        j = j + 1;
//                    }

//                }

//            }
            
//            //try
//            //{

//                int KEnd = Storage1[0].Content.Count+1;
//                wBook.Sheets[Storage1[0].SheetIndex].Cells.EntireColumn.AutoFit();
//                for (int k = 2; k <= KEnd; k++)
//                {
//                Microsoft.Office.Interop.Excel.Range newRng = excelApp.Range[wBook.Sheets[Storage1[0].SheetIndex].Cells[k, 1], wBook.Sheets[Storage1[0].SheetIndex].Cells[k, 30]];
//                    newRng.RowHeight = 20;
//                }
//           // }
//           // catch
//           // {

//           // }
            
            
//        }
//        public void stuff_cell_interior_color(int Color, int CellY, int CellX, int SheetIndex = 1)
//        {
//            wBook.Sheets[SheetIndex].Cells[CellY, CellX].Interior.ColorIndex = Color;

//        }

//        public void stuff_interior_color(int Color,int CellY,int CellXCount,int SheetIndex =1)
//        {
//            for (int i=1;i< CellXCount; i++)
//            {
//                wBook.Sheets[SheetIndex].Cells[CellY, i].Interior.ColorIndex = Color;
//            }
//        }
//        public void stuff_interior_color(int Color, int CellY, int CellXCount, int SheetIndex = 1,int FirstX = 1)
//        {
//            for (int i = FirstX; i < CellXCount; i++)
//            {
//                wBook.Sheets[SheetIndex].Cells[CellY, i].Interior.ColorIndex = Color;
//            }
//        }
//        public void write_title(string[] Title,int Index)
//        {
//            for (int iTitle = 0; iTitle < Title.Length; iTitle++)
//            {
//                wBook.Sheets[Index].Cells[1, iTitle + 1] = Title[iTitle];
//            }

//        }

//        public void append_sheet(string SheetName)
//        {

//            wBook.Sheets.Add(Type.Missing, wBook.Sheets[wBook.Sheets.Count], Type.Missing, Type.Missing); 
//            int Index = wBook.Sheets.Count;
//            wBook.Sheets[Index].Name = SheetName;
//            //Sheets.Add After:= Sheets(Sheets.Count)
//        }

//        public bool wh_exist_sheet(string SheetName)
//        {
//            for (int i=1;i<= wBook.Sheets.Count;i++)
//            {
//                if (wBook.Sheets[i].Name == SheetName)
//                {
//                    return true;
//                }
//            }
//            return false;
//        }

//        #region
//        //public void check_mistake(int TargetCellX, int ErrorCellX)
//        //{
//        //    int iCellY = 2;
//        //    //輸入不同目標位置.檢查不同方式.標色.輸出錯誤字串

//        //    while (wBook.Sheets[1].Cells[iCellY, TargetCellX].Value != null)
//        //    {
//        //        string[] Tgt = new string[2];
//        //        Tgt[0] = wBook.Sheets[1].Cells[iCellY, TargetCellX].Value;
//        //        Tgt[1] = wBook.Sheets[1].Cells[iCellY, TargetCellX].Value;
//        //        string[] ErrStr = new string[2];
//        //        ErrStr[0] = "檔案已經存在無法上傳";
//        //        ErrStr[1] = "此圖片不存在此路徑";
//        //        int[] ErrPtn = new int[2];
//        //        //ErrPtn[0] = "檔案已經存在無法上傳";
//        //        //ErrPtn[1] = "此圖片不存在此路徑";
//        //        //string ErrStr = "";
//        //        //int ErrPtn = 10;

//        //        check_mistake check_mistake1 = new check_mistake(defense.wh_exist);
//        //        check_mistake1 = check_mistake1 += defense.wh_exist;
//        //        int i = 0;
//        //        foreach (check_mistake MyErr in check_mistake1.GetInvocationList())
//        //        {
//        //            if (MyErr.Invoke(Tgt[i], ErrPtn[i], ref ErrStr[i]) == false)
//        //            {
//        //                wBook.Sheets[1].Cells[iCellY, TargetCellX].Interior.ColorIndex = 5;
//        //                wBook.Sheets[1].Cells[iCellY, ErrorCellX].Value = "圖片不存在";
//        //            }
//        //            i = i + 1;
//        //        }


//        //        string PicPath = Convert.ToString(wBook.Sheets[1].Cells[iCellY, TargetCellX].Value);
//        //        if (File.Exists(PicPath) == false)
//        //        {
//        //            wBook.Sheets[1].Cells[iCellY, TargetCellX].Interior.ColorIndex = 5;
//        //            wBook.Sheets[1].Cells[iCellY, ErrorCellX].Value = "圖片不存在";
//        //        }

//        //        iCellY = iCellY + 1;
//        //    }
//        //}
//        ////一個路徑+名稱(建構子)
//        ////預設值
//        #endregion
//        public void test_read_space(Storage[] Storage1)
//        {
//            //先決條件一定要有次序

//            const int READ_COLOR = 15773696;
//            const int CELL_X_ORDER = 1;
//            const int SHEET_FIRST = 1;
//            const int END_TIME = 15;
//            bool WhEnd = false;
//            int SpaceTime = 0;
//            int Temp;
//            //如果catch或是轉0則無法轉
//            int iCellY = 2;
//            while (WhEnd == false)
//            {
//                try
//                {
//                    Temp = Convert.ToInt16(wBook.Sheets[SHEET_FIRST].Cells[iCellY, CELL_X_ORDER].Value);
//                }
//                catch
//                {
//                    Temp = 0;
//                    //MessageBox.Show("error");
//                }

//                //MessageBox.Show(Convert.ToString(wBook.Sheets[SHEET_FIRST].Cells[iCellY, 1].Interior.ColorIndex));
//                if (Temp == 0)
//                {
//                    SpaceTime = SpaceTime + 1;
//                }
//                else if (Temp != 0 && Convert.ToInt32(wBook.Sheets[SHEET_FIRST].Cells[iCellY, 7].Font.Color) == READ_COLOR)
//                {
//                    //MessageBox.Show(Temp.ToString());
//                    //for (int jStuff = 0; jStuff < Storage1.Length; jStuff++)
//                    //{                
//                    //    Storage1[jStuff].Content.Add(Convert.ToString(wBook.Sheets[Storage1[jStuff].SheetIndex].Cells[iCellY, Storage1[jStuff].ReadWriteCellX].Value));
//                    //Storage1[jStuff].Content.Add(wBook.Sheets[Storage1[jStuff].SheetIndex].Cells[iCellY, Storage1[jStuff].ReadWriteCellX].Value);
//                    //}
//                    this.excel_stuff_storage(Storage1, iCellY);
//                    SpaceTime = 0;
//                }
//                if (SpaceTime == END_TIME)
//                {
//                    WhEnd = true;
//                }
//                iCellY = iCellY + 1;
//            }
//        }
//        public void write_cell(int SheetIndex, int CellY, int CellX,string Str)
//        {
//            string TmpStr = wBook.Sheets[SheetIndex].Cells[CellY, CellX].Value;
//            TmpStr = TmpStr + Str;
//            string[] TmpArr = TmpStr.Split(',');
//            string[] RltArr = TmpArr.Distinct().ToArray();
//            string SumStr = "";
//            for (int i=0;i< RltArr.Length;i++)
//            {
//                SumStr = SumStr + RltArr[i] + ",";
//            }
//            wBook.Sheets[SheetIndex].Cells[CellY, CellX].Value = SumStr;
//        }
//        public string read_cell(int SheetIndex,int CellY,int CellX)
//        {
//            return wBook.Sheets[SheetIndex].Cells[CellY, CellX].Value;
//        }
//        public void read_data(Storage Storage1)
//        {
//            int j = 2;
//            //設定若讀取為0則不讀取
//            try
//            {
//                if (Storage1.ReadCellX != 0)
//                {
//                    while (wBook.Sheets[Storage1.SheetIndex].Cells[j, Storage1.ReadCellX].Value != null)
//                    {
//                        Storage1.Content.Add(Convert.ToString(wBook.Sheets[Storage1.SheetIndex].Cells[j, Storage1.ReadCellX].Value));
//                        j = j + 1;
//                    }
//                }
//            }
//            catch
//            {

//            }
            
//        }

//        public void read_data_first_ptn(Storage Storage1)
//        {
//            const int READ_COLOR = 15773696;
//            const int CELL_X_ORDER = 1;
//            const int SHEET_FIRST = 1;
//            const int END_TIME = 15;
//            bool WhEnd = false;
//            int SpaceTime = 0;
//            int Temp;
//            //如果catch或是轉0則無法轉
//            int iCellY = 2;
//            while (WhEnd == false)
//            {
//                try
//                {
//                    Temp = Convert.ToInt16(wBook.Sheets[SHEET_FIRST].Cells[iCellY, CELL_X_ORDER].Value);
//                }
//                catch
//                {
//                    Temp = 0;
//                }


//                if (Temp == 0)
//                {
//                    SpaceTime = SpaceTime + 1;
//                }
//                else if (Temp != 0)
//                {
//                    Storage1.Content.Add(iCellY.ToString());
//                    SpaceTime = 0;
//                }

//                if (SpaceTime == END_TIME)
//                {
//                    WhEnd = true;
//                }

//                iCellY = iCellY + 1;
//            }
//        }
//    }
//}
