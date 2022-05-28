using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Algorithm
{
    public class TestDeny:ITest
    {
        private int test = 2;
        public int Test 
        { 
            get { return this.test; }
            set { this.test = value; } 
        }
        public int TestMethod()
        {
            return 100;

        }
        public int TestMethod(int test)
        {
            return 100;

        }
    }
}
