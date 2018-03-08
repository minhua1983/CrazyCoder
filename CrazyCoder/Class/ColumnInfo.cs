using System;
using System.Collections.Generic;
using System.Text;

namespace CrazyCoder.Class
{
    class ColumnInfo
    {
        string _Name = "";
        string _Type = "nvarchar";
        string _Length = "0";
        string _IsNull = "false";
        string _DefaultValue = "";
        string _Description = "";
        string _Table = "";

        public string Name
        {
            get { return _Name; }
            set { _Name = value; }
        }

        public string Type
        {
            get { return _Type; }
            set { _Type = value; }
        }

        public string Length
        {
            get { return _Length; }
            set { _Length = value; }
        }

        public string IsNull
        {
            get { return _IsNull; }
            set { _IsNull = value; }
        }

        public string DefaultValue
        {
            get { return _DefaultValue; }
            set { _DefaultValue = value; }
        }

        public string Description
        {
            get { return _Description; }
            set { _Description = value; }
        }

        public string Table
        {
            get { return _Table; }
            set { _Table = value; }
        }
    }
}
