using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agbm.Helpers
{
    public static class NumberWrittenTranslator
    {
        private const string ZERO = "No";

        private static readonly string[] _ones = new[] { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine" };
        public static string Translate( this int num )
        {
            if ( num == 0 ) {
                return ZERO;
            }

            return GetLiteral( num, new StringBuilder() );
        }

        private static string GetLiteral( int num, StringBuilder sb )
        {
            if ( num < 10 ) {
                sb.Append( _ones[ num ] );
            }

            return sb.ToString();
        }
    }
}
