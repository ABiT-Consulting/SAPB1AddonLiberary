using ADDONBASE.Attributes;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ADDONBASE.BusinessLogic
{
    public static class SAPB1Helper
    {
        public static void CreateUDT(string className, Company company)
        {
            Type type = Type.GetType(className);
            var tableNameAttribute = (TableNameAttribute)Attribute.GetCustomAttribute(type, typeof(TableNameAttribute));

            if (tableNameAttribute != null)
            {
                UserTablesMD userTable = (UserTablesMD)company.GetBusinessObject(BoObjectTypes.oUserTables);
                userTable.TableName = tableNameAttribute.TableName;
                userTable.TableDescription = tableNameAttribute.Description;
                userTable.TableType = tableNameAttribute.TableType;

                if (userTable.Add() != 0)
                {
                    var ex = new Exception(company.GetLastErrorDescription());
                }

                Marshal.ReleaseComObject(userTable);
            }
        }

        public static void CreateUDF(string className, Company company)
        {
            Type type = Type.GetType(className);
            var tableNameAttribute = (TableNameAttribute)Attribute.GetCustomAttribute(type, typeof(TableNameAttribute));

            PropertyInfo[] properties = type.GetProperties();
            foreach (PropertyInfo property in properties)
            {
                var fieldNameAttribute = (FieldNameAttribute)Attribute.GetCustomAttribute(property, typeof(FieldNameAttribute));

                if (fieldNameAttribute != null)
                {
                    if (!fieldNameAttribute.Name.StartsWith("U_"))
                        continue;
                    var fieldname = fieldNameAttribute.Name.Substring(2);
                    UserFieldsMD userField = (UserFieldsMD)company.GetBusinessObject(BoObjectTypes.oUserFields);
                    userField.TableName = tableNameAttribute.TableName;
                    userField.Name = fieldname;
                    userField.Description = fieldNameAttribute.Description;
                    userField.Type = fieldNameAttribute.FieldType;
                    userField.SubType = fieldNameAttribute.FieldSubType;
                    if (fieldNameAttribute.Size != -1) userField.EditSize = fieldNameAttribute.Size;

                    if (userField.Add() != 0)
                    {
                        var ex = new Exception(company.GetLastErrorDescription());
                    }

                    Marshal.ReleaseComObject(userField);
                }
            }
        }

        public static void CreateUDO(string className, Company company)
        {
            Type type = Type.GetType(className);
            var tableNameAttribute = (TableNameAttribute)Attribute.GetCustomAttribute(type, typeof(TableNameAttribute));
            var udoNameAttribute = (UDONameAttribute)Attribute.GetCustomAttribute(type, typeof(UDONameAttribute));

            if (udoNameAttribute != null)
            {
                UserObjectsMD userObject = (UserObjectsMD)company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                userObject.Code = udoNameAttribute.Name;
                userObject.Name = udoNameAttribute.Name;
                userObject.ObjectType = udoNameAttribute.Type;
                userObject.TableName = tableNameAttribute.TableName;

                PropertyInfo[] properties = type.GetProperties();
                foreach (PropertyInfo property in properties)
                {
                    if (property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
                    {
                        Type childType = property.PropertyType.GetGenericArguments()[0];
                        var childTableNameAttribute = (TableNameAttribute)Attribute.GetCustomAttribute(childType, typeof(TableNameAttribute));

                        if (childTableNameAttribute != null)
                        {
                            userObject.ChildTables.TableName = childTableNameAttribute.TableName;
                            userObject.Add();
                        }
                    }
                }

                if (userObject.Add() != 0)
                {
                    var ex = new Exception(company.GetLastErrorDescription());
                }

                Marshal.ReleaseComObject(userObject);
            }
        }
    }
}
