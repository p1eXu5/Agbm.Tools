using System;

namespace Agbm.NpoiExcel.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class HeaderAttribute : Attribute
    {
        public HeaderAttribute (string header)
        {
            Header = header;
        }

        public string Header { get; set; }

        public override string ToString()
        {
            return Header;
        }
    }
}
