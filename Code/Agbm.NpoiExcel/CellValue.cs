﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Agbm.Helpers;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Agbm.NpoiExcel
{
    public struct CellValue
    {
        public const int DAY_COUNT_NEGATIVE_VARIANCE = -693_595;
        public const int DAY_COUNT_POSITIVE_VARIANCE = 2_899_999;

        private readonly ICell _cell;

        private readonly string _stringValue;
        private double _doubleValue;
        private DateTime _dateTimeValue;

        public CellValue (ICell cell)
        {
            _cell = cell;
            _stringValue = default( string );
            _doubleValue = default( double );
            _dateTimeValue = default( DateTime );
        }

        public double GetDoubleValue ()
        {
            if ( _cell == null ) return _doubleValue;

            if ( _cell.CellType == CellType.Numeric ) {
                return _cell.NumericCellValue;
            }

            if ( _cell.CellType == CellType.String ) {

                var stringValue = _cell.StringCellValue.Trim();

                if ( stringValue.Length > 0 ) {

                    CultureInfo culture = null;

                    culture = CultureInfo.CreateSpecificCulture( "ru-RU" );

                    try {
                        _doubleValue = double.Parse( stringValue, culture );
                    }
                    catch ( FormatException ) {

                        culture = CultureInfo.CreateSpecificCulture( "en-US" );

                        try {
                            _doubleValue = double.Parse( stringValue, culture );
                        }
                        catch ( FormatException ) { }
                    }


                    if ( _doubleValue.Equals( default( double ) ) ) {

                        if ( stringValue.ToUpperInvariant().Equals( "ДА" )
                             || stringValue.ToUpperInvariant().Equals( "YES" ) ) {

                            _doubleValue = 1.0;
                        }
                    }
                }
            }
            else if ( _cell.CellType == CellType.Boolean ) {

                if ( _cell.BooleanCellValue ) {
                    _doubleValue = 1.0;
                }
            }

            return _doubleValue;
        }

        public int GetIntValue ()
        {
            var doubleValue = GetDoubleValue();

            if ( doubleValue > int.MaxValue ) return int.MaxValue;
            if ( doubleValue < int.MinValue ) return int.MinValue;
            return Convert.ToInt32( doubleValue );
        }

        public bool GetBoolValue ()
        {
            if ( _cell == null ) return false;
            if ( _cell.CellType == CellType.Boolean ) return _cell.BooleanCellValue;

            var doubleValue = GetDoubleValue();

            if ( doubleValue.Equals( default( double ) ) ) {
                return false;
            }

            return true;
        }

        public string GetStringValue ()
        {
            if ( _cell == null ) return _stringValue;

            if ( _cell.CellType == CellType.String ) return _cell.StringCellValue;
            if ( _cell.CellType == CellType.Boolean ) return _cell.BooleanCellValue ? "Да" : "Нет";
            if ( _cell.CellType == CellType.Numeric ) return _cell.NumericCellValue.ToString( CultureInfo.InvariantCulture );

            return _stringValue;
        }

        public DateTime GetDateTimeValue ()
        {
            if ( _cell == null || _cell.CellType == CellType.Boolean ) return _dateTimeValue;

            var stringValue = GetStringValue();

            if ( !DateTime.TryParse( stringValue, out _dateTimeValue ) ) {

                var days = GetIntValue();

                if ( days != 0 && days > DAY_COUNT_NEGATIVE_VARIANCE && days <= DAY_COUNT_POSITIVE_VARIANCE ) {

                    _dateTimeValue = new DateTime( 1900, 1, 1 ).AddDays( days - 2 );
                }
            }

            return _dateTimeValue;
        }


        public static implicit operator int (CellValue value)
        {
            return value.GetIntValue();
        }

        public static implicit operator double (CellValue value)
        {
            return value.GetDoubleValue();
        }

        public static implicit operator string (CellValue value)
        {
            return value.GetStringValue();
        }

        public static implicit operator bool (CellValue value)
        {
            return value.GetBoolValue();
        }

        public static implicit operator DateTime (CellValue value)
        {
            return value.GetDateTimeValue();
        }

        public override string ToString()
        {
            return GetStringValue();
        }

        public override bool Equals (object obj)
        {
            if (ReferenceEquals (null, obj)) return false;
            return obj is CellValue other && Equals (other);
        }

        public bool Equals (CellValue other)
        {
            return _stringValue.Equals (other._stringValue);
        }

        public override int GetHashCode()
        {
            return _stringValue.GetHashCode();
        }
    }

}
