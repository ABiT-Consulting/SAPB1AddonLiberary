using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADDONBASE.Attributes
{
    public class UDONameAttribute : Attribute
    {
        // Property to store the UDO name
        public string Name { get; private set; }

        // Property to store the UDO type
        public BoUDOObjType Type { get; private set; }

        // Constructor that takes a name and type parameter
        public UDONameAttribute(string name, BoUDOObjType type)
        {
            Name = name;
            Type = type;
        }
    }
}
