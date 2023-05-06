using SAPbobsCOM;
using System;

namespace ADDONBASE.Attributes
{
    public class FieldNameAttribute : Attribute
    {
        public string Name { get; private set; }
        public string Description { get; private set; }
        public BoFieldTypes FieldType { get; private set; }
        public BoFldSubTypes FieldSubType { get; private set; }
        public int Size { get; private set; }

        public FieldNameAttribute(string fieldName, string fieldDescription, BoFieldTypes fieldType, BoFldSubTypes fieldSubType, int size = -1)
        {
            Name = fieldName;
            Description = fieldDescription;
            FieldType = fieldType;
            FieldSubType = fieldSubType;
            Size = size;
        }
    }
}
