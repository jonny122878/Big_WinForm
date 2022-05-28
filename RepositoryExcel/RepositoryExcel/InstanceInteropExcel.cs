using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
namespace RepositoryExcel
{



    public class InstanceInteropExcel : InteropExcel, IReadAndWriteForExcel
    {
        public InstanceInteropExcel(string file, bool WhVisible = true) : base(file, WhVisible)
        {

        }
        public void WriteToTable<TEntity>(int y, int x, List<TEntity> data, Func<TEntity, List<IExcelInfo>> map) where TEntity : class, new()
        {
            base.WriteToTable<TEntity>(y,x, data, map);
        }

        public void WriteToCell<TEntity>(IExcelInfo data)
        {
            base.WriteToCell<TEntity>(data);
        }


        public void Save(bool isClose = true)
        {
            base.Save(isClose);
        }
    }
}
