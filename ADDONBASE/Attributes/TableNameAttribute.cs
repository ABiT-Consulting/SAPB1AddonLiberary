using SAPbobsCOM;
using System;

namespace ADDONBASE.Attributes
{
    public class TableNameAttribute : Attribute
    {
        public string TableName { get; private set; }
        public string Description { get; private set; }
        public BoUTBTableType TableType { get; private set; }

        public TableNameAttribute(string tableName, string description, BoUTBTableType tableType)
        {
            if (tableName.StartsWith("@"))
                tableName = tableName.Substring(1);

            TableName = tableName;
            Description = description;
            TableType = tableType;
        }
    }
}
