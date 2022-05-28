using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Algorithm
{
    public interface ITest
    {
        int Test { get; set; }
        int TestMethod();
        int TestMethod(int test);
    }
}
