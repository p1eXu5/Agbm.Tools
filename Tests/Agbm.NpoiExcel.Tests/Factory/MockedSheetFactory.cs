using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Moq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;

namespace Agbm.NpoiExcel.Tests.Factory
{
    public class MockedSheetFactory
    {
        public static ISheet EmptySheet => new XSSFSheet();

        /// <summary>
        /// Return test cases.
        /// </summary>
        /// <param name="firstColumn"></param>
        /// <param name="firsrtRow"></param>
        /// <returns></returns>
        public static IEnumerable HeaderTestCases (int firstColumn, int firsrtRow)
        {
            // 1
            var area = new[,]
                {
                    { 0, 0, 0, 1 },
                    { 0, 0, 1, 0 },
                    { 0, 1, 0, 0 },
                    { 1, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .SetName($"Test case #1 ({firstColumn}, {firsrtRow});");

            // 2
            area = new[,]
                {
                    { 1, 0, 0, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 0, 0, 1 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .SetName($"Test case #2 ({firstColumn}, {firsrtRow});");

            // 3
            area = new[,]
                {
                    { 1, 0, 0, 0 },
                    { 1, 1, 0, 0 },
                    { 1, 0, 1, 1 },
                    { 1, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .SetName($"Test case #3 ({firstColumn}, {firsrtRow});");

            // 4
            area = new[,]
                {
                    { 1, 1, 1, 1 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 1, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .SetName($"Test case #4 ({firstColumn}, {firsrtRow});");

            // 5
            area = new[,]
                {
                    { 0, 0, 0, 1 },
                    { 0, 1, 0, 1 },
                    { 1, 0, 1, 1 },
                    { 0, 0, 0, 1 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .SetName($"Test case #5 ({firstColumn}, {firsrtRow});");

            // 6
            area = new[,]
                {
                    { 0, 1, 0, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 1, 1, 1, 1 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .SetName($"Test case #6 ({firstColumn}, {firsrtRow});");

            // 7
            area = new[,]
                {
                    { 0, 0, 0, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .SetName($"Test case #7 ({firstColumn}, {firsrtRow});");

            // 8
            area = new[,]
                {
                    { 0, 0, 0, 0 },
                    { 0, 1, 1, 0 },
                    { 0, 0, 0, 0 },
                    { 0, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                    .SetName($"Test case #8 ({firstColumn}, {firsrtRow});");

            // 9
            area = new[,]
                {
                    { 0, 0, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 0, 0, 0 },
                    { 0, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .SetName($"Test case #9 ({firstColumn}, {firsrtRow});");
        }

        /// <summary>
        /// Return test cases.
        /// </summary>
        /// <param name="firstColumn"></param>
        /// <param name="firsrtRow"></param>
        /// <returns></returns>
        public static IEnumerable HeaderTestCasesWithReturns (int firstColumn, int firsrtRow)
        {
            // 1
            var area = new[,]
                {
                    { 0, 0, 0, 1 },
                    { 0, 0, 1, 0 },
                    { 0, 1, 0, 0 },
                    { 1, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .Returns (new[]{"", "", "", "TESTSTRING" })
                                .SetName($"Test case #1 ({firstColumn}, {firsrtRow});");

            // 2
            area = new[,]
                {
                    { 1, 0, 1, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 0, 0, 1 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .Returns(new[] { "TESTSTRING", "", "TESTSTRING", "" })
                                .SetName($"Test case #2 ({firstColumn}, {firsrtRow});");

            // 4
            area = new[,]
                {
                    { 1, 1, 1, 1 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 1, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .Returns(new[] { "TESTSTRING", "TESTSTRING", "TESTSTRING", "TESTSTRING" })
                                .SetName($"Test case #4 ({firstColumn}, {firsrtRow});");


            // 7
            area = new[,]
            {
                { 0, 0, 0, 0 },
                { 0, 1, 0, 0 },
                { 0, 1, 0, 0 },
                { 0, 0, 0, 0 }
            };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                 .Returns(new[] { "TESTSTRING" })
                                 .SetName($"Test case #7 ({firstColumn}, {firsrtRow});");

            // 8
            area = new[,]
                {
                    { 0, 0, 0, 0 },
                    { 0, 1, 1, 0 },
                    { 0, 0, 0, 0 },
                    { 0, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                 .Returns(new[] { "TESTSTRING", "TESTSTRING" })
                                 .SetName($"Test case #8 ({firstColumn}, {firsrtRow});");

            // 9
            area = new[,]
                {
                    { 0, 0, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 0, 0, 0 },
                    { 0, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                 .Returns(new[] { "TESTSTRING" })
                                 .SetName($"Test case #9 ({firstColumn}, {firsrtRow});");
        }

        /// <summary>
        /// Return test cases.
        /// </summary>
        /// <param name="firstColumn"></param>
        /// <param name="firsrtRow"></param>
        /// <returns></returns>
        public static IEnumerable RowColumnCountsTestCases(int firstColumn, int firsrtRow)
        {
            // 1
            var area = new[,]
                {
                    { 0, 0, 0, 1 },
                    { 0, 0, 1, 0 },
                    { 0, 1, 0, 0 },
                    { 1, 0, 0, 0 }
                };
            yield return new TestCaseData(GetMockedSheet((short)firstColumn, firsrtRow, area))
                                .Returns((3, 4))
                                .SetName($"Test case #1 ({firstColumn}, {firsrtRow});");

            // 2
            area = new[ , ]
                {
                    { 1, 0, 0, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 0, 0, 1 }
                };
            yield return new TestCaseData( GetMockedSheet( ( short )firstColumn, firsrtRow, area ) )
                                .Returns( (3, 4) )
                                .SetName( $"Test case #2 ({firstColumn}, {firsrtRow});" );

            // 3
            area = new[ , ]
                {
                    { 1, 0, 0, 0 },
                    { 1, 1, 0, 0 },
                    { 1, 0, 1, 1 },
                    { 1, 0, 0, 0 }
                };
            yield return new TestCaseData( GetMockedSheet( ( short )firstColumn, firsrtRow, area ) )
                                .Returns( (3, 4) )
                                .SetName( $"Test case #3 ({firstColumn}, {firsrtRow});" );

            // 4
            area = new[ , ]
                {
                    { 1, 1, 1, 1 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 1, 0, 0 }
                };
            yield return new TestCaseData( GetMockedSheet( ( short )firstColumn, firsrtRow, area ) )
                                .Returns( (3, 4) )
                                .SetName( $"Test case #4 ({firstColumn}, {firsrtRow});" ).SetCategory( "Headers" );

            // 5
            area = new[ , ]
                {
                    { 0, 0, 0, 1 },
                    { 0, 1, 0, 1 },
                    { 1, 0, 1, 1 },
                    { 0, 0, 0, 1 }
                };
            yield return new TestCaseData( GetMockedSheet( ( short )firstColumn, firsrtRow, area ) )
                                .Returns( (3, 4) )
                                .SetName( $"Test case #5 ({firstColumn}, {firsrtRow});" );

            // 6
            area = new[ , ]
                {
                    { 0, 1, 0, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 1, 1, 1, 1 }
                };
            yield return new TestCaseData( GetMockedSheet( ( short )firstColumn, firsrtRow, area ) )
                                .Returns( (3, 4) )
                                .SetName( $"Test case #6 ({firstColumn}, {firsrtRow});" );

            // 7
            area = new[ , ]
                {
                    { 0, 0, 0, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 1, 0, 0 },
                    { 0, 0, 0, 0 }
                };
            yield return new TestCaseData( GetMockedSheet( ( short )firstColumn, firsrtRow, area ) )
                                .Returns( (1, 1) )
                                .SetName( $"Test case #7 ({firstColumn}, {firsrtRow});" );

            // 8
            area = new[ , ]
                {
                    { 0, 0, 0, 0 },
                    { 0, 1, 1, 0 },
                    { 0, 0, 0, 0 },
                    { 0, 0, 0, 0 }
                };
            yield return new TestCaseData( GetMockedSheet( ( short )firstColumn, firsrtRow, area ) )
                                    .Returns( (0, 2) )
                                    .SetName( $"Test case #8 ({firstColumn}, {firsrtRow});" );

            // 9
            area = new[ , ]
                {
                    { 0, 0, 0, 0 },
                    { 0, 0, 1, 0 },
                    { 0, 0, 0, 0 },
                    { 0, 0, 0, 0 }
                };
            yield return new TestCaseData( GetMockedSheet( ( short )firstColumn, firsrtRow, area ) )
                                .Returns( (0, 1) )
                                .SetName( $"Test case #9 ({firstColumn}, {firsrtRow});" );
        }

        public static ISheet GetMockedSheet ( short firstColumn, int firsrtRow, int[,] area )
        {
            var rowsMap = new Dictionary<int, HashSet<int>>();

            for ( int j = 0; j < area.GetLength( 0 ); j++ ) {
                for ( short i = 0; i < area.GetLength( 1 ); ++i ) {

                    if ( area[ j, i ] != 0 ) {

                        var rowNumber = j + firsrtRow;

                        if ( !rowsMap.Keys.Contains( rowNumber ) ) {
                            rowsMap[ rowNumber ] = new HashSet<int>();
                        }

                        rowsMap[ rowNumber ].Add( i + firstColumn );
                    }
                }
            }

            return GetMockedSheet( rowsMap );
        }

        public static ISheet GetMockedSheet ( string[,] tableData )
        {
            if ( tableData == null ) throw new ArgumentNullException( nameof( tableData ), "Headers cannot be null." );
            if ( tableData.Length <= 0 ) return EmptySheet;

            var columnMap = new HashSet< int >( Enumerable.Range( 0, tableData.GetLength( 1 ) ) );

            var rowsMap = new Dictionary< int, HashSet< int > > ();

            for ( int i = 0; i < tableData.GetLength( 0 ); ++i ) {
                rowsMap[ i ] = columnMap;
            }

            return GetMockedSheet( rowsMap, tableData );
        }

        public static ISheet GetMockedSheet ( Type inType, ITypeRepository typeRepo, object[] data )
        {
            var cellMap = new Dictionary< int, HashSet< int > >();
            var properties = typeRepo.GetPropertyNames( inType ).ToArray();
            var columnCount = properties.Length;

            // headers ar random row
            cellMap[ 113 ] = new HashSet< int >( Enumerable.Range( 11, columnCount ) );
            cellMap[ 114 ] = new HashSet< int >( Enumerable.Range( 11, columnCount ) );

            object[,] cellData = new object[ 2, columnCount ];

            // load headers
            for ( int i = 0; i < columnCount; i++ ) {
                cellData[ 0, i ] = properties[ i ];
                cellData[ 1, i ] = data[ i ];
            }

            return GetMockedSheet( cellMap, cellData );
        }

        /// <summary>
        /// Creates mocked sheet from cell map ( [rowNum] = columnSet ) and cell data.
        /// </summary>
        /// <param name="cellMap"><see cref="Dictionary{TKey,TValue}"/></param>
        /// <param name="data"></param>
        /// <returns></returns>
        private static ISheet GetMockedSheet ( Dictionary< int, HashSet< int > > cellMap, object[,] data )
        {
            // см. типы передаваемых параметров, а не возвращаемых!
            var sheetMock = new Mock< ISheet >();
            var firstRow = cellMap.Keys.Min();
            var lastRow = cellMap.Keys.Max();
            var firstColumn = cellMap[ firstRow ].Min();

            sheetMock.Setup( s => s.FirstRowNum ).Returns( firstRow );
            sheetMock.Setup( s => s.LastRowNum ).Returns( lastRow );

            sheetMock.Setup( s => s.GetRow( It.Is< int >( r => cellMap.Keys.Contains( r ) ) ) )
                     .Returns( ( int r ) =>
                               {
                                   var rowMock = new Mock< IRow >();
                                   rowMock.Setup( s => s.FirstCellNum ).Returns( ( short )cellMap[ r ].Min() );
                                   rowMock.Setup( s => s.LastCellNum ).Returns( ( short )(cellMap[ r ].Max() + 1) );
                                   rowMock.Setup( row => row.GetCell( It.Is< int >( c => cellMap[ r ].Contains( c ) ) ) )
                                          .Returns( ( int c ) => GetMockedCell( data, r - firstRow, c - firstColumn ) );
                                   return rowMock.Object;
                               } );

            return sheetMock.Object;
        }

        private static ISheet GetMockedSheet ( Dictionary< int, HashSet< int > > cellMap )
        {
            // см. типы передаваемых параметров, а не возвращаемых!
            var sheetMock = new Mock< ISheet >();

            sheetMock.Setup( s => s.FirstRowNum ).Returns( cellMap.Keys.Min() );
            sheetMock.Setup( s => s.LastRowNum ).Returns( cellMap.Keys.Max() );

            sheetMock.Setup( s => s.GetRow( It.Is< int >( r => cellMap.Keys.Contains( r ) ) ) )
                     .Returns( ( int r ) =>
                               {
                                   var rowMock = new Mock< IRow >();
                                   rowMock.Setup( s => s.FirstCellNum ).Returns( ( short )cellMap[ r ].Min() );
                                   rowMock.Setup( s => s.LastCellNum ).Returns( ( short )(cellMap[ r ].Max() + 1) );
                                   rowMock.Setup( row => row.GetCell( It.Is< int >( c => cellMap[ r ].Contains( c ) ) ) )
                                          .Returns( ReturnMockedStringCell() );
                                   return rowMock.Object;
                               } );

            return sheetMock.Object;
        }

        private static ICell GetMockedCell ( object[,] data, int row, int column )
        {
            return GetMockedCell( (dynamic)data[ row, column ] );
        }

        private static ICell GetMockedCell ( object value )
        {
            var stringCellMock = new Mock<ICell>();

            stringCellMock.Setup( c => c.CellType ).Returns( CellType.String );
            stringCellMock.Setup( c => c.StringCellValue ).Returns( value.ToString() );

            return stringCellMock.Object;
        }

        private static ICell GetMockedCell ( string value )
        {
            var stringCellMock = new Mock<ICell>();

            stringCellMock.Setup( c => c.CellType ).Returns( CellType.String );
            stringCellMock.Setup( c => c.StringCellValue ).Returns( value );

            return stringCellMock.Object;
        }

        private static ICell GetMockedCell ( double value )
        {
            var stringCellMock = new Mock<ICell>();

            stringCellMock.Setup( c => c.CellType ).Returns( CellType.Numeric );
            stringCellMock.Setup( c => c.NumericCellValue ).Returns( value );

            return stringCellMock.Object;
        }

        private static ICell GetMockedCell ( bool value )
        {
            var stringCellMock = new Mock<ICell>();

            stringCellMock.Setup( c => c.CellType ).Returns( CellType.Boolean );
            stringCellMock.Setup( c => c.BooleanCellValue ).Returns( value );

            return stringCellMock.Object;
        }

        private static ICell ReturnMockedStringCell ()
        {
            var stringCellMock = new Mock<ICell>();

            stringCellMock.Setup( c => c.CellType ).Returns( CellType.String );
            stringCellMock.Setup( c => c.StringCellValue ).Returns( "Test string" );

            return stringCellMock.Object;
        }
    }
}
