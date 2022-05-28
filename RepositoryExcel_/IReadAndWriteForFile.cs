using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RepositoryExcel
{

    //Formula 每個類下面的合計 或是特定指定要分母除

    //interface基本上一個就夠了不用到2個
    //因為excel部分結構算是固定的所以用抽象類去取代
    
    //IReadAndWriteForFile 讀取檔案這個抽象層次因txt 、xml結構完全不同
    //但行為相似所以建立interface
    
    //instance Null Object Test Object
    public interface IReadAndWriteForExcel
    {
        #region 這個部分先暫且忽略不考慮..以後如果要擴充只讀取紅色..或其他規則時
        ///// <summary>
        ///// 將資料檔案轉成物件. (停止點為整個 row 為 empty 或者為 null)
        ///// </summary>
        ///// <param name="filePath"> 讀取路徑. </param>
        ///// <param name="fileName"> 檔名+副檔名. </param>
        ///// <returns> 包含資料的自定義物件. </returns>
        //List<TEntity> ReadForFile<TEntity>(string filePath, string fileName, int startRow, System.Drawing.Color color)
        //    where TEntity : class, new();
        #endregion


        /// <summary>
        /// 文件檔改成在abstract class去實作
        /// 因為結構在這個層次上是不相同的，可以不用在層次是去抽象
        /// 參數過多不適用
        /// 將自定義物件寫入至特定資料型檔案
        /// 這裡部分ExcelInfo之後納入txt做抽換
        /// 探討問題參數的順序
        /// </summary>
        /// <param name="entity"> 含有資料的自定義物件. </param>
        /// <param name="fullpath"> 樣板檔案的完整路徑. </param>
        /// <param name="fullPathExport"> 產出檔案的完整路徑. </param>
        /// <param name="tilteModels"> Excel 內文標題.(會放在最上方) </param>
        void WriteToTable<TEntity>(int y, int x,List<TEntity> data,Func<TEntity, List<IExcelInfo>> map)
            where TEntity : class,new();
        void WriteToRange<TEntity>(int y1, int x1, int y2, int x2, IExcelInfo data);
        void WriteToCell<TEntity>(int y, int x, IExcelInfo data);

        void Save(bool isClose = true);

    }
}
