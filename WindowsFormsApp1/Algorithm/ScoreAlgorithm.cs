using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WindowsFormsApp1.Algorithm.Models;
using WindowsFormsApp1.ViewModels;
using RepositoryExcel.ViewModels;
namespace WindowsFormsApp1.Algorithm
{
    public class ScoreAlgorithm
    {
        private ScoreAlgorithmModel _scoreAlgorithmModel;
        private ScoreViewModel _scoreViewModel;
        public ScoreAlgorithm(ScoreAlgorithmModel scoreAlgorithmModel,
                              ScoreViewModel scoreViewModel)
        {
            _scoreAlgorithmModel = scoreAlgorithmModel;
            _scoreViewModel = scoreViewModel;
            //實作條件式
        }     

       
        //# 計算高分群調整比例
        //        AdjustRate_H = rounddown(Top1_to/rank1)

        //        # 判斷是否為高分群調整例外狀況
        //        rank6 = score[5]
        //        if (HighRanK > 5) and(rank6, AdjustRate_H > 89.5) :
        //            AdjustRate_H = rounddown(89.5/rank6)

        //        # 計算中分群調整比例
        //        AdjustRate_M = rounddown(((HopeAverage* people)-(H* AdjustRate_H)-L)/M)

        //        #中分群調整比例限額計算
        //        if AdjustRate_M > AdjustRate_H:

        //總共三種狀況：高分群、高中分群、高中分群加限額
        //條件判別：AdjustRate、AdjustRate_M、AdjustRate_H 來做區分演算法
        public Tuple<ExcelState,TableAttrExcelModel, List<ScoreResultExcelModel>> GetExcelResult(List<ScoreExcelModel> scoreExcelModels)
        {
            
            //AdjustRate
            decimal adjustRate = (
                            (_scoreViewModel.AvgOffset * _scoreAlgorithmModel.PeopleQty) -
                            _scoreAlgorithmModel.MiddenGroupSum -
                            _scoreAlgorithmModel.LowGroupSum) / _scoreAlgorithmModel.HighGroupSum;
            decimal adjustRateR2 = Math.Round(adjustRate, 2);

            //AdjustRate_H = rounddown(Top1_to/rank1)
            decimal adjustRate_H = _scoreViewModel.FirstOffset / _scoreAlgorithmModel.Avg_R_Top;
            //if (HighRanK > 5) and (rank6,AdjustRate_H > 89.5):
            if (_scoreViewModel.HighGroup > 5 && _scoreAlgorithmModel.Avg_R_H > 89.5m)
            {
                adjustRate_H = 89.5m / _scoreAlgorithmModel.Avg_R_H;
            }
            //AdjustRate_M
            //AdjustRate_M = rounddown(((HopeAverage*people)-(H*AdjustRate_H)-L)/M)
            decimal adjustRate_H2 = Math.Round(adjustRate_H, 2);

            decimal adjustRate_M = (_scoreViewModel.AvgOffset * _scoreAlgorithmModel.PeopleQty) -
                                    ((_scoreAlgorithmModel.HighGroupSum * adjustRate_H) -
                                    _scoreAlgorithmModel.LowGroupSum)
                                    / _scoreAlgorithmModel.MiddenGroupSum;
            decimal adjustRate_M2 = Math.Round(adjustRate_M, 2);
            if(adjustRate_M2 > adjustRate_H)
            {
                adjustRate_M2 = adjustRate_H;
            }

            //output:state
            #region WHERE enum狀態
            List<Func<ExcelState>> conditions = new List<Func<ExcelState>>
            {
                () => { return (adjustRateR2 < (100m / _scoreAlgorithmModel.Avg_R_Top))?
                                ExcelState.HIGH: ExcelState.NONE;
                      },
                () => { return (!(adjustRate_M > adjustRate_H))?
                                ExcelState.HIGH_MIDDLE: ExcelState.NONE;
                      },
                () => { return (adjustRate_M > adjustRate_H)?
                                ExcelState.HIGH_MIDDLE_LIMIT: ExcelState.NONE;
                      },
            };
            var condition = conditions.FirstOrDefault(r => r.Invoke() != ExcelState.NONE);
            var state = (condition != null) ? condition.Invoke() : ExcelState.NONE;
            #endregion
            Console.WriteLine("");
            

            //output:tableAttr
            #region WHERE 表格相關資訊
            Dictionary<ExcelState, TableAttrExcelModel> dictAttrs = new Dictionary<ExcelState, TableAttrExcelModel>();
            dictAttrs.Add(ExcelState.HIGH, new TableAttrExcelModel { HighScope = adjustRateR2 });
            dictAttrs.Add(ExcelState.HIGH_MIDDLE, new TableAttrExcelModel
            {
                HighScope = adjustRate_H2,
                MiddleScope = adjustRate_M2
            });
            dictAttrs.Add(ExcelState.HIGH_MIDDLE_LIMIT, new TableAttrExcelModel
            {
                HighScope = adjustRate_H2,
                MiddleScope = adjustRate_M2,
                //LessSum =

            });
            var tableAttr = dictAttrs.FirstOrDefault(r => r.Key == state).Value;
            if (tableAttr == null)
            {
                tableAttr = new TableAttrExcelModel();
            }
            #endregion
            Console.WriteLine("");



            //將python score_adjust完成
            Dictionary<ExcelState, Func<List<decimal>>> dicts = new Dictionary<ExcelState, Func<List<decimal>>>();
            //dicts.Add(ExcelState.HIGH,() => );

            //只調整Average_R Excel第八欄 去*要調整倍率
            //其它不變完整輸出上
            List<decimal> average_Rs = new List<decimal>();

            var results = scoreExcelModels.Zip(average_Rs,
                (inner, outer) =>
            new ScoreResultExcelModel
            {
                No = inner.No,
                Student = inner.Student,
                Midterm = inner.Midterm,
                Final = inner.Final,
                In_Class = inner.In_Class,

                Homework = inner.Homework,
                Average = inner.Average,
                Average_R = outer.ToString()
            }).ToList();

            return new Tuple<ExcelState, TableAttrExcelModel, List<ScoreResultExcelModel>>(state, tableAttr, results);
        }
    }
}
//AdjustRate = ((HopeAverage* people)-M-L)/H
//  AdjustRate = rounddown(AdjustRate)

//    # 判斷是否為方法一的例外狀況
//rank1 = score[0]
//    if AdjustRate< 100/rank1 : 