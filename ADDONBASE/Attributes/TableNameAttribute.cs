using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADDONBASE.Attributes
{
    public class TableNameAttribute:Attribute
    {
        string TABLE_NAME= "";
        SAPbobsCOM.BoUTBTableType BoUTBTableType;
        string Description = "";
        public TableNameAttribute(string tablename,string Description ,SAPbobsCOM.BoUTBTableType boUTBTableType)
        {
            if (tablename.StartsWith("@"))
                tablename =  tablename.Substring(1);
            this.TABLE_NAME = tablename;
            this.BoUTBTableType = boUTBTableType;
            this.Description = Description;
        }
    }
}
