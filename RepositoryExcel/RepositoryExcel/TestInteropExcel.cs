//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Microsoft.Office.Interop.Excel;
//using System.Reflection;
//namespace RepositoryExcel
//{
//    //由於保護層級關係要做測試私有方法
//    public class TestInteropExcel : InteropExcel, IReadAndWriteForExcel
//    {
//      public TestInteropExcel(string file, bool WhVisible = true):base(file, WhVisible)
//        {


//        }

//        public List<List<IExcelInfo>> GetexcelInfies<TEntity>(List<TEntity> data, Func<TEntity, List<IExcelInfo>> map)
//        {
//            return base.GetexcelInfies<TEntity>(data, map);
//        }
//        public void Loadcoodrinates(List<List<IExcelInfo>> excelInfies, int y, int x)
//        {
//            base.Loadcoodrinates(excelInfies,y,x);
//        }
//        public void LoadTitle<TEntity>(List<List<IExcelInfo>> excelInfies, int y, int x)
//        {
//            base.LoadTitle<TEntity>(excelInfies, y, x);
//        }
        
//        public List<Tuple<int, int, IExcelInfo>> GetcellEnumerable()
//        {
//            return this.GetcellEnumerable();
//        }
//        public void WriteInfo(List<Tuple<int, int, IExcelInfo>> cellEnumerable)
//        {
//            base.WriteInfo(cellEnumerable);
//        }
//        public void PrintBorder(List<Tuple<int, int, IExcelInfo>> cellEnumerable)
//        {
//            base.PrintBorder(cellEnumerable);
//        }

//        public void WriteToCell<TEntity>(int y, int x, IExcelInfo data)
//        {
//            throw new NotImplementedException();
//        }

//        public void WriteToRange<TEntity>(int y1, int x1, int y2, int x2, IExcelInfo data)
//        {
//            throw new NotImplementedException();
//        }

//        public void WriteToTable<TEntity>(int y, int x, List<TEntity> data, Func<TEntity, List<IExcelInfo>> map) where TEntity : class, new()
//        {
//            throw new NotImplementedException();
//        }
//    }
//}
