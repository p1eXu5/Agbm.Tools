using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agbm.NpoiExcel.Tests.IntegrationalTests.Cases
{
    class DoubleCases
    {
        private static object[] NumericStringCell =
        {
            new object[] { "0.0", 0.0 },
            new object[] { "0", 0.0 },
            new object[] { "-9876543.0123456789", -9876543.0123456789 },
            new object[] { "9876543.0123456789", 9876543.0123456789 },
            new object[] { " 9876543", 9876543.0 },
            new object[] { "-9876543 ", -9876543.0 },
        };
    }
}
