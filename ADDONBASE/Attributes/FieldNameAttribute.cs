using SAPbobsCOM;
using System;

namespace ADDONBASE.Attributes
{
    public class FieldNameAttribute : Attribute
    {
        private string FieldName;
        private BoFieldTypes boFieldTypes;
        private BoFldSubTypes boFldSubTypes;
        private string FieldDescription;
        private int size;
        public FieldNameAttribute(string FieldName,string FieldDescription, BoFieldTypes boFieldTypes, BoFldSubTypes boFldSubTypes,int size = -1)
        {
            this.FieldName = FieldName;
            this.boFieldTypes = boFieldTypes;
            this.boFldSubTypes = boFldSubTypes;
            this.size = size;
            this.FieldDescription = FieldDescription;
        }
    }
}