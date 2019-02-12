using System;

namespace Agbm.NpoiExcel.Tests.Factory
{
    public class TypeExcelFactory
    {
        public static Type EmptyClass => typeof (Empty);
        public static Type ClassWithParameterizedCtor => typeof (ParameterizedCtor);


        class ParameterizedCtor
        {
            public ParameterizedCtor (int foo) { }
        }
    }

    public class Empty { }
}
