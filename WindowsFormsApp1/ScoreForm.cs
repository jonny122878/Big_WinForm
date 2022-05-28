using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Data.OleDb;
using LinqToExcel;
using RepositoryExcel.ViewModels;
using WindowsFormsApp1.ViewModels;
using WindowsFormsApp1.Algorithm.Models;
using RepositoryExcel;
using WindowsFormsApp1.Algorithm;
namespace WindowsFormsApp1
{
    public partial class ScoreForm : Form
    {
        //第三個參數為label值、四參為text預設值
        protected List<Tuple<Label, TextBox, string, string>> _View;

        //Load來源
        protected ScoreViewModel scoreViewModel;
        protected List<ScoreExcelModel> excelModels;
        protected ScoreAlgorithmModel scoreAlgorithmModel;
        //運算端
        protected ScoreAlgorithm scoreAlgorithm;
        protected IReadAndWriteForExcel excelContext;
        //產出
        protected List<ScoreResultExcelModel> excelData;

        public ScoreForm()
        {
            InitializeComponent();
            //將View端給初始化並賦值
            this._View = new List<Tuple<Label, TextBox, string, string>>
            {
                new Tuple<Label, TextBox, string, string>(this.label1,this.textBox1,"名次於多少前為高分群","10"),
                new Tuple<Label, TextBox, string, string>(this.label2,this.textBox2,"名次於多少後為低分群","25"),
                new Tuple<Label, TextBox, string, string>(this.label3,this.textBox3,"將第一名提升到多少分","10"),
                new Tuple<Label, TextBox, string, string>(this.label4,this.textBox4,"將班平均調整到多少","80"),
                new Tuple<Label, TextBox, string, string>(this.label5,this.textBox5,"匯入成績Excel檔",@"D:\會計系統\文件\打分數樣品\打分數樣品\打分數樣品相關料\電腦試算成績.xlsx"),
            };
            this._View.ForEach(item =>
            {
                item.Item1.Text = item.Item3;
                item.Item2.Text = item.Item4;
            });
            //this.button1_Click(null,null);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //參數output:ScoreViewModel
            #region View端參數載入ViewModel 
            //驗證點可能錯誤
            //表單數字轉型失敗
            //excel 檔案不存在、無sheet、資料載入不出來
            scoreViewModel = new ScoreViewModel();
            var arr = this._View.Select(r => r.Item2.Text).ToArray();
            scoreViewModel.HighGroup = Convert.ToInt32(arr[0]);
            scoreViewModel.LowGroup =  Convert.ToInt32(arr[1]);
            scoreViewModel.FirstOffset = new Decimal(Convert.ToDouble(arr[2]));
            scoreViewModel.AvgOffset = new Decimal(Convert.ToDouble(arr[3]));
            scoreViewModel.Excel = arr[4];
            //Load 參數
            #endregion
            Console.Write("");

            //output:excelModels
            //參數output:ScoreAlgorithmModel
            #region Load Excel資料
            //驗證點：檔案不存在、工作表不存在、被其他人使用、無資料、自帶資料格式
            //資料數值轉型
            var excelFile = scoreViewModel.Excel;
            var sheet = "成績試算";
            var excelReadContext = new ExcelQueryFactory(excelFile);
            excelModels = excelReadContext.Worksheet<ScoreExcelModel>(sheet).ToList();
            Console.Write("");
            var orderAvgRs = excelModels.Select(r => new Decimal(Convert.ToDouble(r.Average_R))).
                OrderBy(r => r).ToList();
            var orderDescAvgRs = orderAvgRs.OrderByDescending(r => r).ToList();
            #region 另種判別演算法高中低總分
            //var scoreAlgorithmModel = new ScoreAlgorithmModel
            //{
            //    PeopleQty = excelModels.Count,
            //    Avg_R = orderAvgRs.First(),
            //    Sum = orderAvgRs.Sum(),
            //    LowGroupSum = orderAvgRs.Where(r => r <= scoreViewModel.LowGroup).Sum(),
            //    MiddenGroupSum = orderAvgRs.Where(r => r < scoreViewModel.HighGroup &&
            //    r > scoreViewModel.LowGroup).Sum(),
            //    HighGroupSum = orderAvgRs.Where(r => r >= scoreViewModel.HighGroup).Sum(),
            //};
            #endregion
            //驗算數量
            //var LowGroups = orderDescAvgRs.Skip(scoreViewModel.LowGroup).Count();
            //var MiddenGroups = orderDescAvgRs.Skip(scoreViewModel.HighGroup).Take(scoreViewModel.LowGroup - scoreViewModel.HighGroup).Count();
            //var HighGroups = orderDescAvgRs.Take(scoreViewModel.HighGroup).Count();
            scoreAlgorithmModel = new ScoreAlgorithmModel
            {
                PeopleQty = excelModels.Count,
                Avg_R_Top = orderAvgRs.First(),
                Avg_R_H = orderAvgRs.Skip(5).First(),
                Sum = orderDescAvgRs.Sum(),
                LowGroupSum = orderDescAvgRs.Skip(scoreViewModel.LowGroup).Sum(),
                MiddenGroupSum = orderDescAvgRs.Skip(scoreViewModel.HighGroup).Take(scoreViewModel.LowGroup - scoreViewModel.HighGroup).Sum(),
                HighGroupSum = orderDescAvgRs.Take(scoreViewModel.HighGroup).Sum(),
            };
            #endregion

            Console.Write("");

            //output:result、tableAttr(產生出表格狀態、表格相關資訊、內容)
            //Load algorithm:scoreAlgorithm、excelContext
            #region Load algorithm

            scoreAlgorithm = new ScoreAlgorithm(scoreAlgorithmModel, scoreViewModel);
            var excelResultFile = @"D:\會計系統\文件\打分數樣品\打分數樣品\打分數樣品相關料\產出結果.xlsx";
            excelContext = new InstanceInteropExcel(excelResultFile);
            var result = scoreAlgorithm.GetExcelResult(excelModels);
            var tableAttr = result.Item2;
            #endregion
            Console.Write("");

            #region excel打印表身各種狀態模型

            

            excelContext.WriteToCell<ScoreResultExcelModel>(
                   new ExcelInfoCell
                   {
                       Y = 5,
                       X = 12,
                       Value = "高分群調整比例：" + tableAttr.HighScope.ToString()
                   });
            excelContext.WriteToCell<ScoreResultExcelModel>(
               new ExcelInfoCell
               {
                   Y = 5,
                   X = 12,
                   Value = "中分群調整比例：" + tableAttr.MiddleScope.ToString()
               });
            if (tableAttr.LessSum != 0m)
            {
                excelContext.WriteToCell<ScoreResultExcelModel>(
                new ExcelInfoCell
                {
                    Y = 5,
                    X = 12,
                    Value = string.Format(@"全班總分尚需加{0}分，才能達至預期平均", tableAttr.LessSum.ToString())
                }); ;
            }

            #endregion
            Console.Write("");


            #region excel各種狀態都要打印表格資訊
            excelContext.WriteToCell<ScoreResultExcelModel>(
                   new ExcelInfoCell
                   {
                       Y = 1,
                       X = 12,
                       Value = "Average_R平均："
                   });

            excelContext.WriteToCell<ScoreResultExcelModel>(
                new ExcelInfoCell
                {
                    Y = 2,
                    X = 12,
                    Value = "Adjusted平均："
                });

            excelContext.WriteToCell<ScoreResultExcelModel>(
                new ExcelInfoCell
                {
                    Y = 4,
                    X = 12,
                    Value = "調整狀況："
                });

            excelContext.WriteToCell<ScoreResultExcelModel>(
                new ExcelInfoCell
                {
                    Y = 1,
                    X = 13,
                    Value = "=ROUND(AVERAGE($H:$H),2)"
                });

            excelContext.WriteToCell<ScoreResultExcelModel>(
                new ExcelInfoCell
                {
                    Y = 2,
                    X = 13,
                    Value = "=ROUND(AVERAGE($I:$I),2)"
                });

            #endregion
            Console.Write("");

            #region excel打印表格內容

            #region old python

            //        ws['L1'] = "Average_R平均："
            //ws['L2'] = "Adjusted平均："
            //ws['M1'] = r"=ROUND(AVERAGE($H:$H),2)"
            //ws['M2'] = r"=ROUND(AVERAGE($I:$I),2)"


            //        if method1 == True:
            //    ws['L5'] = "高分群調整比例：" + str(AdjustRate)
            //    ws['L6'] = "中分群調整比例：0"

            //elif(method2 == True) and(enough == True):
            //    ws['L5'] = "高分群調整比例：" + str(AdjustRate_H)
            //    ws['L6'] = "中分群調整比例：" + str(AdjustRate_M)

            //elif(method2 == True) and(enough == False):
            //    ws['L5'] = "高分群調整比例：" + str(AdjustRate_H)
            //    ws['L6'] = "中分群調整比例：" + str(AdjustRate_M)


            //    NeedToAdd = HopeAverage * people - rank['Adjusted'].sum()
            //    ws['L7'] = "全班總分尚需加" + str(NeedToAdd) + "分，才能達至預期平均"

            //else:
            //    ws['L5'] = "數據有誤，請洽維修人員" 
            #endregion           

            Func<ScoreResultExcelModel, List<IExcelInfo>> map = r =>
            {

                return new List<IExcelInfo>() 
                { 
                    new ExcelInfoCell{ Value = r.No,Y = 2,X = 1,},
                    new ExcelInfoCell{ Value = r.Student,Y = 2,X = 2,},
                    new ExcelInfoCell{ Value = r.Midterm,Y = 2,X = 3,},
                    new ExcelInfoCell{ Value = r.Final,Y = 2,X = 4,},
                    new ExcelInfoCell{ Value = r.In_Class,Y = 2,X = 5,},

                    new ExcelInfoCell{ Value = r.Homework,Y = 2,X = 6,},
                    new ExcelInfoCell{ Value = r.Average,Y = 2,X = 7,},
                    new ExcelInfoCell{ Value = r.Average_R,Y = 2,X = 8,},
                    new ExcelInfoCell{ Value = r.Adjusted,Y = 2,X = 9,},

                };
            };
            excelContext.WriteToTable<ScoreResultExcelModel>(1, 1, excelData, map);
            excelContext.WriteOverride();
            #endregion
            MessageBox.Show("ok");
            Console.Write("");
        }
    }
}
