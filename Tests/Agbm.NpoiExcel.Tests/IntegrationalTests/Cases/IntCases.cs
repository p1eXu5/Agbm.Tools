using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agbm.NpoiExcel.Tests.IntegrationalTests.Cases
{
    class IntCases
    {
        private static object[] NumericCell =
        {
            new object[] { 0.0, 0 },
            new object[] { -9876543.0123456789, -9876543 },
            new object[] { 9876543.0123456789, 9876543 }
        };

        private static object[] NumericStringCell =
        {
            new object[] { "0.0", 0 },
            new object[] { "0", 0 },
            new object[] { "-9876543.0123456789", -9876543 },
            new object[] { "9876543.0123456789", 9876543 },
            new object[] { " 9876543", 9876543 },
            new object[] { "-9876543 ", -9876543 },
        };

        private static object[] BooleanStringCell =
        {
            new object[] { "Да", 1 },
            new object[] { "да", 1 },
            new object[] { "дА", 1 },
            new object[] { "ДА", 1 },
            new object[] { "Yes", 1 },
        };

        private static string[] StringCell =
        {
            "d0.0",
            "r0",
            "-9876543.012+3456789",
            "9876543.01 23456789",
            "d9876543",
            "-987/6543",
            "No",
            "nO",
            "Нет",
        };

        private static object[] DateTimeCell =
        {
            new object[] { "11.12.2015", 42349 },
            new object[] { "01.01.1900", 1 },
            new object[] { "25.07.2541 0:34:12", 234328 },
            new object[] { "13.03.1346", -1 },

        };

        private static string[] NullableString = 
        {
            "null",
            "NULL  ",
            "  nUlL",
            "      ",
            ""
        };
    }
}
