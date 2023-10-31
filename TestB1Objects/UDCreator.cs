 
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text; 
using System.Xml.Linq; 

namespace TestB1Objects
{
    public class UDTCreator
    {
        SAPbobsCOM.Company com; 
        public UDTCreator(SAPbobsCOM.Company com)
        {
            this.com = com;
        } 
        public String message { get; set; }
        public bool CreateTable(string TableName, string TableDesc, SAPbobsCOM.BoUTBTableType TableType)
        {
            bool functionReturnValue = false;
            functionReturnValue = false;
            long v_RetVal = 0;
            int v_ErrCode = 0;
            string v_ErrMsg = "";
            try
            {
                if (!this.TableExists(TableName))
                {
                    SAPbobsCOM.UserTablesMD v_UserTableMD = default(SAPbobsCOM.UserTablesMD);
                    // app.StatusBar.SetText("Creating Table " + TableName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    v_UserTableMD = (SAPbobsCOM.UserTablesMD)com.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                    v_UserTableMD.TableName = TableName;
                    v_UserTableMD.TableDescription = TableDesc;
                    v_UserTableMD.TableType = TableType;

                    v_RetVal = v_UserTableMD.Add();
                    if (v_RetVal != 0)
                    {

                        com.GetLastError(out v_ErrCode, out v_ErrMsg);
                        if (v_ErrCode != -2035)
                        {
                            throw new Exception("Failed to Create Table " + TableDesc + v_ErrCode + " " + v_ErrMsg);//, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD);
                            v_UserTableMD = null;
                            return false;
                        }
                    }
                    else
                    {
                        // app.StatusBar.SetText("[" + TableName + "] - " + TableDesc + " Created Successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD);
                        v_UserTableMD = null;
                        return true;
                    }
                }
                else
                {
                    GC.Collect();
                    return false;
                }
            }
            catch (Exception ex)
            {
                //if (System.Diagnostics.Debugger.IsAttached)
                //{
                //    System.Diagnostics.Debugger.Break();

                //}
                throw new Exception(ex.Message + " @ " + ex.Source);
            }
            return functionReturnValue;
        }
        public bool TableExists(string TableName)
        {
            SAPbobsCOM.UserTablesMD oTables = default(SAPbobsCOM.UserTablesMD);
            bool oFlag = false;
            oTables = (SAPbobsCOM.UserTablesMD)com.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            oFlag = oTables.GetByKey(TableName);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oTables);
            return oFlag;
        }
        public bool CreateUserFields(string TableName, string FieldName, string FieldDescription, SAPbobsCOM.BoFieldTypes type, long size = 0, SAPbobsCOM.BoFldSubTypes subType = SAPbobsCOM.BoFldSubTypes.st_None, string LinkedTable = "", Int32 EditSize = 0, string byDefaultValue = "")
        {
            bool returnValue = false;
            try
            {

                if (TableName.StartsWith("@") == true)
                {
                    if (!this.ColumnExists(TableName, FieldName))
                    {
                        SAPbobsCOM.UserFieldsMD v_UserField = default(SAPbobsCOM.UserFieldsMD);
                        v_UserField = (SAPbobsCOM.UserFieldsMD)com.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                        try
                        {
                            v_UserField.TableName = TableName;
                            v_UserField.Name = FieldName;
                            v_UserField.Description = FieldDescription;
                            v_UserField.Type = type;

                            if (type != SAPbobsCOM.BoFieldTypes.db_Date)
                            {
                                if (size != 0)
                                {
                                    v_UserField.Size = (int)size;
                                }
                                if (EditSize != 0)
                                {
                                    v_UserField.EditSize = EditSize;
                                }
                            }
                            if (subType != SAPbobsCOM.BoFldSubTypes.st_None)
                            {
                                v_UserField.SubType = subType;
                            }
                            int v_RetVal;
                            int v_ErrCode = 0;
                            string v_ErrMsg = "";
                            if (!string.IsNullOrEmpty(LinkedTable))
                                v_UserField.LinkedTable = LinkedTable;
                            v_RetVal = v_UserField.Add();

                            if (v_RetVal != 0)
                            {
                                com.GetLastError(out v_ErrCode, out v_ErrMsg);
                                if (v_ErrCode != -2035)
                                {
                                    throw new Exception("Failed to add UserField masterid" + v_ErrCode + " " + v_ErrMsg);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                                    v_UserField = null;
                                    returnValue = false;
                                }

                            }
                            else
                            {
                                //app.StatusBar.SetText("[" + TableName + "] - " + FieldDescription + " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                                v_UserField = null;
                                returnValue = true;
                            }
                        }finally
                        {
                            if(v_UserField != null)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                        }
                    }
                    else
                    {
                        returnValue = false;
                    }
                }

                if (TableName.StartsWith("@") == false)
                {

                    if (!this.UDFExists(TableName, FieldName))
                    {
                        SAPbobsCOM.UserFieldsMD v_UserField = default(SAPbobsCOM.UserFieldsMD);
                        v_UserField = (SAPbobsCOM.UserFieldsMD)com.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                        try
                        {
                            v_UserField.TableName = TableName;
                            v_UserField.Name = FieldName;
                            v_UserField.Description = FieldDescription;
                            v_UserField.Type = type;
                            if (type != SAPbobsCOM.BoFieldTypes.db_Date)
                            {
                                if (size != 0)
                                {
                                    v_UserField.Size = (int)size;
                                }
                            }
                            if (!string.IsNullOrEmpty(byDefaultValue))
                            {
                                v_UserField.DefaultValue = byDefaultValue;
                            }
                            if (subType != SAPbobsCOM.BoFldSubTypes.st_None)
                            {
                                v_UserField.SubType = subType;
                            }
                            int v_RetVal = 0;
                            int v_ErrCode = 0;
                            string v_ErrMsg = "";
                            if (!string.IsNullOrEmpty(LinkedTable))
                                v_UserField.LinkedTable = LinkedTable;
                            v_RetVal = v_UserField.Add();
                            if (v_RetVal != 0)
                            {
                                com.GetLastError(out v_ErrCode, out v_ErrMsg);

                                //app.StatusBar.SetText("Failed to add UserField " + FieldDescription + " - " + v_ErrCode + " " + v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                                v_UserField = null;
                                returnValue = false;


                            }
                            else
                            {
                                // app.StatusBar.SetText(" & TableName & - " + FieldDescription + " added successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                                v_UserField = null;
                                returnValue = true;
                            }
                        }finally
                        {
                            if (v_UserField != null)
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                        }
                    }
                    else
                    {
                        returnValue = false;
                    }
                }
            }
            catch (Exception ex)
            {
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    System.Diagnostics.Debugger.Break();

                }
                //   app.StatusBar.SetText(ex.Message);
                returnValue = false;
            }
            return returnValue;
        }
        string ColumnExists_1
        {
            get
            {
                return "Select 1 from \"CUFD\" Where \"TableID\"='{0}' and \"AliasID\"='{1}'";
            }
        }
        public bool ColumnExists(string TableName, string FieldID)
        {

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)com.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            bool oFlag = true;
            var s = string.Format(ColumnExists_1, TableName.Trim(), FieldID.Trim());
            rs.DoQuery(s);//"Select 1 from [CUFD] Where TableID='" + TableName.Trim() + "' and AliasID='" + FieldID.Trim() + "'");
            if (rs.EoF)
                oFlag = false;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;
            GC.Collect();
            return oFlag;

        }
        public bool UDFExists(string TableName, string FieldID)
        {
            bool oFlag = true;
            try
            {
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)com.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                var s = string.Format(ColumnExists_1, TableName.Trim(), FieldID.Trim());
                rs.DoQuery(s);//"Select 1 from [CUFD] Where TableID='" + TableName.Trim() + "' and AliasID='" + FieldID.Trim() + "'");

                //    rs.DoQuery("Select 1 from [CUFD] Where TableID='" + TableName.Trim() + "' and AliasID='" + FieldID.Trim() + "'");
                if (rs.EoF)
                    oFlag = false;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                rs = null;
                GC.Collect();

            }
            catch (Exception ex)
            {
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    System.Diagnostics.Debugger.Break();

                }
                //   app.StatusBar.SetText(ex.Message);
                oFlag = false;

            }
            return oFlag;
        }

    }
    public class UDOCreator
    {
        SAPbobsCOM.Company ocompany;
        public UDOCreator(SAPbobsCOM.Company ocompany)
        {
            this.ocompany = ocompany;
        }
        public bool UDOExists(string code)
        {
            GC.Collect();
            SAPbobsCOM.UserObjectsMD v_UDOMD = default(SAPbobsCOM.UserObjectsMD);
            bool v_ReturnCode = false;
            v_UDOMD = (SAPbobsCOM.UserObjectsMD)ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            v_ReturnCode = v_UDOMD.GetByKey(code);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD);
            v_UDOMD = null;
            return v_ReturnCode;
        }

        public bool RegisterUDO(string UDOCode, string UDOName, SAPbobsCOM.BoUDOObjType UDOType, string[,] FindField, string UDOHTableName, string UDODTableName = "", SAPbobsCOM.BoYesNoEnum CanCancel = global::SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum CanClose = global::SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum CanDelete = global::SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum CanFind = global::SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum CanLog = global::SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum ManageSeries = global::SAPbobsCOM.BoYesNoEnum.tYES, string ChildTable = "", string ChildTable1 = "", SAPbobsCOM.BoYesNoEnum LogOption = SAPbobsCOM.BoYesNoEnum.tNO, string ChildTable2 = "", string menueid = "", string menuecaption = "", SAPbobsCOM.BoYesNoEnum candefaultForm = SAPbobsCOM.BoYesNoEnum.tYES)
        {

            bool functionReturnValue = false;
            bool ActionSuccess = false;
            try
            {
                functionReturnValue = false;
                SAPbobsCOM.UserObjectsMD v_udoMD = default(SAPbobsCOM.UserObjectsMD);
                v_udoMD = this.ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;

                v_udoMD.CanCancel = CanCancel;// SAPbobsCOM.BoYesNoEnum.tYES;
                v_udoMD.CanClose = CanClose;// SAPbobsCOM.BoYesNoEnum.tYES;
                //    v_udoMD.CanCreateDefaultForm = candefaultForm ;
                v_udoMD.CanDelete = CanDelete; // SAPbobsCOM.BoYesNoEnum.tYES;
                v_udoMD.CanFind = CanFind;// SAPbobsCOM.BoYesNoEnum.tYES;
                v_udoMD.CanLog = CanLog;// SAPbobsCOM.BoYesNoEnum.tYES;
                //v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tyes
                v_udoMD.ManageSeries = ManageSeries;

                v_udoMD.Code = UDOCode;
                v_udoMD.Name = UDOName;
                v_udoMD.TableName = UDOHTableName;


                if (!string.IsNullOrEmpty(UDODTableName))
                {
                    v_udoMD.ChildTables.TableName = UDODTableName;
                    v_udoMD.ChildTables.Add();
                }

                if (!string.IsNullOrEmpty(ChildTable))
                {
                    v_udoMD.ChildTables.TableName = ChildTable;
                    v_udoMD.ChildTables.Add();
                }
                if (!string.IsNullOrEmpty(ChildTable1))
                {
                    v_udoMD.ChildTables.TableName = ChildTable1;
                    v_udoMD.ChildTables.Add();
                }
                if (!string.IsNullOrEmpty(ChildTable2))
                {
                    v_udoMD.ChildTables.TableName = ChildTable2;
                    v_udoMD.ChildTables.Add();
                }

                if (LogOption == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    v_udoMD.LogTableName = "A" + UDOHTableName;
                    v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                v_udoMD.ObjectType = UDOType;
                for (Int16 i = 0; i <= FindField.GetLength(0) - 1; i++)
                {
                    //   if (i > 0)

                    v_udoMD.FindColumns.ColumnAlias = FindField[i, 0];
                    v_udoMD.FindColumns.ColumnDescription = FindField[i, 1];
                    v_udoMD.FindColumns.Add();
                }
                //if (menueid != "")
                //{
                //    v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;

                //    v_udoMD.MenuItem = SAPbobsCOM.BoYesNoEnum.tYES; 

                //    v_udoMD.MenuCaption = menuecaption;
                //    v_udoMD.FatherMenuID =int.Parse ( menueid);
                //    v_udoMD.Position = 0;
                //    v_udoMD.MenuUID = UDOName;
                //}

                if (v_udoMD.Add() == 0)
                {
                    functionReturnValue = true;
                    if (ocompany.InTransaction)
                        ocompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                else
                {
                    functionReturnValue = false;
                    throw new Exception("Failed to Register UDO >" + UDOCode + ">" + UDOName + " >" + ocompany.GetLastErrorDescription());

                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD);
                v_udoMD = null;
                GC.Collect();
                if (menueid != "")
                {
                    try
                    {
                        SAPbobsCOM.UserObjectsMD udo = ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
                        udo.CanYearTransfer = candefaultForm;
                        // Get UDO 
                        udo.GetByKey(UDOCode);
                        udo.FormColumns.FormColumnAlias = "DocEntry";
                        //udo.FormColumns.FormColumnDescription = "docentry";
                        udo.Add();
                        // Set UDO to have a menu 
                        udo.MenuItem = SAPbobsCOM.BoYesNoEnum.tYES;
                        udo.MenuCaption = UDOName;

                        // Set father and position of menu item. 
                        udo.FatherMenuID = int.Parse(menueid); // Business Partners menu UID 
                        udo.Position = 1;

                        // Set UDO menu UID 
                        udo.MenuUID = UDOCode + "m";

                        // Update UDO to have the new menu item 
                        var done = udo.Update();
                        if (done != 0)
                        {
                            int errCode = 0;
                            string errMsg = "";
                            ocompany.GetLastError(out errCode, out errMsg);

                            // Application.SBO_Application.SetStatusBarMessage(errMsg);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (System.Diagnostics.Debugger.IsAttached)
                        {
                            System.Diagnostics.Debugger.Break();

                        }
                    }
                }
                if (ActionSuccess == false & ocompany.InTransaction)
                    ocompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }
            catch (Exception ex)
            {
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    System.Diagnostics.Debugger.Break();

                }
                if (ocompany.InTransaction)
                    ocompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                throw new Exception(string.Format("{0}-{1}", ex.Message, ex.StackTrace));
            }

            return functionReturnValue;
        }

    }

    public class udcreator
    {
        SAPbobsCOM.Company company;
        public udcreator(SAPbobsCOM.Company company)
        {
            this.company = company;

            //var id = ut.Add();
        } 
        public void createTablesfromXML(string xmlfile)
        {
            XDocument document = XDocument.Parse(System.IO.File.ReadAllText(xmlfile).Replace(@"xmlns=""http://udt.org""", ""));
            XElement tablename = document.Element("TableName");
            XElement obj = document.Element("Object");

            if (document.Element("BOM").Element("BO").Element("AdmInfo").Element("Object").Value == "153")
            {
                var rows = document.Element("BOM").Element("BO").Element("OUTB").Elements("row");
                var total = rows.Count();
                var count = 0;

                foreach (var row in rows)
                {
                    bool returnValue;
                    count++;
                    var uts = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables) as SAPbobsCOM.UserTables;
                   
                        SAPbobsCOM.UserTablesMD v_UserTableMD = default(SAPbobsCOM.UserTablesMD);
                        // app.StatusBar.SetText("Creating Table " + TableName + " ...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        v_UserTableMD = (SAPbobsCOM.UserTablesMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                        v_UserTableMD.TableName = row.Element("TableName").Value;
                        v_UserTableMD.TableDescription = row.Element("Descr").Value;
                        v_UserTableMD.TableType = (SAPbobsCOM.BoUTBTableType)int.Parse(row.Element("ObjectType").Value);
                        v_UserTableMD.Archivable = (row.Element("Archivable").Value == "N") ? SAPbobsCOM.BoYesNoEnum.tNO : SAPbobsCOM.BoYesNoEnum.tYES;

                        var v_RetVal = v_UserTableMD.Add();
                        if (v_RetVal != 0)
                        {
                            int v_ErrCode = 0;
                            string v_ErrMsg = "";
                            company.GetLastError(out v_ErrCode, out v_ErrMsg);
                            if (v_ErrCode != -2035)
                            {
                                throw new Exception("Failed to Create Table " + row.Element("TableName").Value + v_ErrCode + " " + v_ErrMsg);
                            }
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD);
                        v_UserTableMD = null;
                        GC.Collect(); 

                }
            }
        }
        public bool ColumnExists(string TableName, string FieldID)
        {
            
            bool oFlag = true;
            var company = this.company;

            var s = string.Format(ColumnExists_1, TableName.Trim(), FieldID.Trim());
            // todo: do the following logic with company recordset rather than dbhelper
            var recset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(s);
            if (recset.RecordCount == 0)
                oFlag = false;
            else
                oFlag = true; 

            Marshal.ReleaseComObject(recset);
            return oFlag;

        }
        string ColumnExists_1
        {
            get
            {
                return "Select 1 from \"CUFD\" Where \"TableID\"='{0}' and \"AliasID\"='{1}'";
                //return ConfigurationManager.AppSettings["ColumnExists_1"];
            }
        }
        public bool UDFExists(string TableName, string FieldID)
        {
            bool oFlag = true;
            try
            {
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                var s = string.Format(ColumnExists_1, TableName, FieldID);

                rs.DoQuery(s);//"Select 1 from [CUFD] Where TableID='" + TableName.Trim() + "' and AliasID='" + FieldID.Trim() + "'");
                if (rs.EoF)
                    oFlag = false;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                rs = null;
                GC.Collect();

            }
            catch (Exception ex)
            {
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    System.Diagnostics.Debugger.Break();

                }
                //   app.StatusBar.SetText(ex.Message);
                oFlag = false;

            }
            return oFlag;
        }
        public void createUDFFromXML(string xmlfile)
        {
            XDocument document = XDocument.Parse(System.IO.File.ReadAllText(xmlfile).Replace(@"xmlns=""http://udf.org""", ""));
            XElement tablename = document.Element("TableName");
            XElement obj = document.Element("Object");
            if (document.Element("BOM").Element("BO").Element("AdmInfo").Element("Object").Value == "152")
            {
                var rows = document.Element("BOM").Element("BO").Element("CUFD").Elements("row").Where(x => x.Attribute("resolved").Value == "N");
                var total = rows.Count();
                var count = 0;

                foreach (var row in rows)
                {


                    #region Declarations
                    string TableName = string.Empty;
                    string FieldName = string.Empty;
                    string FieldDescription = string.Empty;
                    string type = string.Empty;
                    SAPbobsCOM.BoFieldTypes Type = default(SAPbobsCOM.BoFieldTypes);
                    int size = 0;
                    int EditSize = 0;
                    SAPbobsCOM.BoFldSubTypes subType = default(SAPbobsCOM.BoFldSubTypes);
                    string LinkedTable = string.Empty;
                    string byDefaultValue = string.Empty;
                    bool returnValue;
                    #endregion
                    #region Initialization
                    try
                    {
                        TableName = row.Element("TableName")?.Value.Trim();
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        FieldName = row.Element("Name")?.Value.Trim();
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        FieldDescription = row.Element("Description")?.Value.Trim();
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        type = row.Element("Type")?.Value.ToLower();
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        Type = (type == "a") ? SAPbobsCOM.BoFieldTypes.db_Alpha : (type == "d") ? SAPbobsCOM.BoFieldTypes.db_Date : (type == "f") ? SAPbobsCOM.BoFieldTypes.db_Float : (type == "m") ? SAPbobsCOM.BoFieldTypes.db_Memo : SAPbobsCOM.BoFieldTypes.db_Numeric;
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        size = int.Parse(row.Element("Size")?.Value);
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        EditSize = int.Parse(row.Element("EditSize")?.Value);
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        subType = (SAPbobsCOM.BoFldSubTypes)int.Parse(row.Element("SubType")?.Value);
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        LinkedTable = row.Element("LinkedTable")?.Value;
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    try
                    {
                        byDefaultValue = row.Element("DefaultValue")?.Value;
                    }
                    catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                    #endregion
                    #region FOR UDF

                    if (!this.ColumnExists(TableName, FieldName))
                    {
                         
                            SAPbobsCOM.UserFieldsMD v_UserField = (SAPbobsCOM.UserFieldsMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                        try {  
                            v_UserField.TableName = TableName;
                            v_UserField.Name = FieldName;
                            v_UserField.Description = FieldDescription;
                            v_UserField.Type = Type;
                            v_UserField.SubType = SAPbobsCOM.BoFldSubTypes.st_None;
                            if (!string.IsNullOrEmpty(byDefaultValue))
                            {
                                v_UserField.DefaultValue = byDefaultValue;
                            }
                            try
                            {
                                var elements = row.Elements("ValidValues");
                                foreach (var v in elements)
                                {
                                    v_UserField.ValidValues.Value = v.Element("Value").Value;
                                    v_UserField.ValidValues.Description = v.Element("Description").Value;
                                    v_UserField.ValidValues.Add();
                                }
                            }
                            catch (Exception ex) { Program.oapp.SetStatusBarMessage(ex.Message); }
                            if (Type != SAPbobsCOM.BoFieldTypes.db_Date)
                            {
                                if (size != 0)
                                {
                                    v_UserField.Size = (int)size;
                                }
                                if (EditSize != 0)
                                {
                                    v_UserField.EditSize = EditSize;
                                }
                            }
                            if (subType != SAPbobsCOM.BoFldSubTypes.st_None)
                            {
                                v_UserField.SubType = subType;
                            }
                            int v_RetVal;
                            int v_ErrCode = 0;
                            string v_ErrMsg = "";
                            if (!string.IsNullOrEmpty(LinkedTable))
                                v_UserField.LinkedTable = LinkedTable;
                            v_RetVal = v_UserField.Add();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                            v_UserField = null;
                            if (v_RetVal != 0)
                            {
                                company.GetLastError(out v_ErrCode, out v_ErrMsg);
                                if (v_ErrCode != -2035)
                                {
                                    row.Attribute("resolved").SetValue(v_ErrMsg);
                                    returnValue = false;
                                }
                                Program.oapp.SetStatusBarMessage(v_ErrMsg);

                            }
                            else
                            {
                                row.Attribute("resolved").SetValue("Y");
                                returnValue = true;
                                Program.oapp.SetStatusBarMessage("Column " + FieldName + " added to table " + TableName);

                            }
                        }
                        finally
                        {
                            if (v_UserField != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField);
                                v_UserField = null;
                            }
                        }
                    }
                    else
                    {
                        returnValue = false;
                        Program.oapp.SetStatusBarMessage("Column " + FieldName + " already exists in table " + TableName);
                    }
                    #endregion 
                    Program.oapp.SetStatusBarMessage("Total Rows " + total + " Processed " + count);
                    count++;

                }
                GC.Collect();
            }
            document.Save(xmlfile);


        }
        public bool createUDOFromXML(string xmlfile)
        {
            XDocument document = XDocument.Parse(System.IO.File.ReadAllText(xmlfile).Replace("http://udo.org", ""));
            XElement tablename = document.Element("TableName");
            XElement obj = document.Element("Object");


            var created = false;

            if (document.Element("BOM").Element("BO").Element("AdmInfo").Element("Object").Value == "206")
            {
                var rows = document.Element("BOM").Element("BO").Element("OUDO").Elements("row").Where(x => x.Attribute("resolved").Value == "N");
                var total = rows.Count();
                var count = 0;

                foreach (var row in rows)
                {
                    count++;
                    #region create udo for OHEM
                    try
                    {

                        // UDOCreator creator = new UDOCreator(company);


                        string[,] FindField = new string[,] { { "Code", "Code" }, { "Name", "Name" } };

                        SAPbobsCOM.UserObjectsMD v_udoMD = default(SAPbobsCOM.UserObjectsMD);
                        v_udoMD = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
                        var v_ReturnCode = v_udoMD.GetByKey(row.Element("Code").Value.Trim());
                        if (!v_ReturnCode)
                        {
                            //   v_udoMD.GetType().GetProperty(row.Name.LocalName).SetValue(v_udoMD, row.Value,null);
                            #region create
                            v_udoMD.Code = row.Element("Code").Value.Trim();
                            v_udoMD.Name = row.Element("Name").Value.Trim();
                            v_udoMD.ObjectType = (SAPbobsCOM.BoUDOObjType)int.Parse(row.Element("ObjectType").Value.Trim());
                            v_udoMD.TableName = row.Element("TableName").Value.Trim();
                            v_udoMD.LogTableName = row.Element("LogTableName").Value.Trim();
                            v_udoMD.ManageSeries = (SAPbobsCOM.BoYesNoEnum)((row.Element("ManageSeries").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            v_udoMD.CanFind = (SAPbobsCOM.BoYesNoEnum)((row.Element("CanFind").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            try
                            {
                                if (row.Element("EnableEnhancedForm") != null)
                                    v_udoMD.EnableEnhancedForm = (SAPbobsCOM.BoYesNoEnum)((row.Element("EnableEnhancedForm").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            }
                            catch (Exception Ex) { Program.oapp.SetStatusBarMessage(Ex.Message); }
                            try
                            {
                                if (row.Element("RebuildEnhancedForm") != null)
                                    v_udoMD.RebuildEnhancedForm = (SAPbobsCOM.BoYesNoEnum)((row.Element("RebuildEnhancedForm").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            }
                            catch (Exception ex) { v_udoMD.RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO; }
                            v_udoMD.CanDelete = (SAPbobsCOM.BoYesNoEnum)((row.Element("CanDelete").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            v_udoMD.CanClose = (SAPbobsCOM.BoYesNoEnum)((row.Element("CanClose").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            v_udoMD.CanCancel = (SAPbobsCOM.BoYesNoEnum)((row.Element("CanCancel").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            v_udoMD.CanLog = (SAPbobsCOM.BoYesNoEnum)((row.Element("CanLog").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            v_udoMD.CanYearTransfer = (SAPbobsCOM.BoYesNoEnum)((row.Element("CanYearTransfer").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            v_udoMD.FormColumns.FormColumnAlias = row.Element("FormColumnAlias").Value.Trim();// "DocEntry";
                            v_udoMD.FormColumns.FormColumnDescription = row.Element("FormColumnDescription").Value.Trim();// "DocEntry";
                            v_udoMD.CanCreateDefaultForm = (SAPbobsCOM.BoYesNoEnum)((row.Element("CanCreateDefaultForm").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                            v_udoMD.UseUniqueFormType = SAPbobsCOM.BoYesNoEnum.tNO;

                            try
                            {
                                if (row.Element("MenuItem") != null)
                                {
                                    v_udoMD.MenuItem = (SAPbobsCOM.BoYesNoEnum)((row.Element("MenuItem").Value.Trim().ToLower().Equals("n")) ? 0 : 1);
                                    if (v_udoMD.MenuItem == SAPbobsCOM.BoYesNoEnum.tYES)
                                    {
                                        v_udoMD.MenuCaption = row.Element("MenuCaption").Value;
                                        v_udoMD.FatherMenuID = Convert.ToInt32(row.Element("FatherMenuID").Value);
                                        v_udoMD.Position = Convert.ToInt32(row.Element("Position").Value);
                                    }
                                }
                            }
                            catch (Exception Ex) { Program.oapp.SetStatusBarMessage(Ex.Message); }
                            if (v_udoMD.CanFind == SAPbobsCOM.BoYesNoEnum.tYES)
                            {
                                if (v_udoMD.ObjectType == SAPbobsCOM.BoUDOObjType.boud_Document)
                                { FindField = new string[,] { { "DocEntry", "DocEntry" } }; }
                                for (Int16 i = 0; i <= FindField.GetLength(0) - 1; i++)
                                {
                                    //   if (i > 0)

                                    v_udoMD.FindColumns.ColumnAlias = FindField[i, 0];
                                    v_udoMD.FindColumns.ColumnDescription = FindField[i, 1];
                                    v_udoMD.FindColumns.Add();
                                }

                            }

                            var FormColumns = row.Elements("FormColumns");
                            foreach (var v in FormColumns)
                            {
                                v_udoMD.FormColumns.FormColumnAlias = v.Element("FormColumnAlias").Value.Trim();
                                v_udoMD.FormColumns.FormColumnDescription = v.Element("FormColumnDescription").Value.Trim();
                                v_udoMD.FormColumns.Add();
                            }
                            var childrenTable = row.Elements("ChildTables");
                            foreach (var v in childrenTable)
                            {
                                v_udoMD.ChildTables.TableName = v.Element("TableName").Value.Trim();
                                v_udoMD.ChildTables.LogTableName = v.Element("LogTableName").Value.Trim();

                                v_udoMD.ChildTables.Add();
                            }
                            #endregion
                            //try
                            //{
                            //    var path = @"srf\" + v_udoMD.Code + ".srf";
                            //    if (System .IO.File .Exists (path)){
                            //    var xml1 = System.IO.File.ReadAllText(path );

                            //    v_udoMD.FormSRF = xml1;
                            //    }
                            //}
                            //catch(Exception ex) { ex.AppendInLogFile(); }
                            // v_udoMD.FormSRF = row.Element("FormSRF").Value;
                            var id = v_udoMD.Add();
                            if (id != 0)
                            {
                                row.Attribute("resolved").Value = company.GetLastErrorDescription();
                            }
                            else
                            {
                                try
                                {
                                    var key = company.GetNewObjectKey();
                                    var Value = row.Element("CanNewForm").Value.Trim();
                                    string query = string.Format("update  OUDO set \"CanNewForm\"='{1}' where \"Code\" = '{0}'", key, Value);

                                    var oRecset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                    oRecset.DoQuery(query);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecset);
                                    oRecset = null;
                                    GC.Collect();
                                }
                                catch
                                { }
                                created = true;
                            }
                            Console.WriteLine(row.Attribute("resolved").Value);
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD); }

                            catch (Exception ex) { }
                            v_udoMD = null;
                            //    udocreator.RegisterUDO("sbsudo", "sbsudo", SAPbobsCOM.BoUDOObjType.boud_Document, FindField, "osbsch", "sbsch1","", "", SAPbobsCOM.BoYesNoEnum.tYES,"","2048");
                        }
                    }
                    catch (Exception ex1)
                    {
                        Program.oapp.SetStatusBarMessage(ex1.Message);
                        //  Application.SBO_Application.SetStatusBarMessage(ex1.Message);
                    }
                    #endregion


                }

                document.Save(xmlfile);


            }
            return created;
        }

        public List<String> getUDONames(String xmlfile)
        {
            XDocument document = XDocument.Parse(System.IO.File.ReadAllText(xmlfile).Replace("http://udo.org", ""));
            XElement tablename = document.Element("TableName");
            XElement obj = document.Element("Object");

            List<String> strs = new List<string>();

            var created = false;

            if (document.Element("BOM").Element("BO").Element("AdmInfo").Element("Object").Value == "206")
            {
                var rows = document.Element("BOM").Element("BO").Element("OUDO").Elements("row").Where(x => x.Attribute("resolved").Value == "N");
                foreach (var row in rows)
                {
                    strs.Add(row.Element("Code").Value.Trim());


                }



            }
            return strs;
        }

        internal void createFMSFromXML(string file)
        {
            var oCompany = this.company;
            var oFS = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObjectFromXML(file, 0);//,SAPbobsCOM.BoObjectTypes.oFormattedSearches)));

            // oFS.Browser.ReadXml(file, 0);
            var id = oFS.Add();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oFS);

            if (id != 0)
            {

                throw new Exception(oCompany.GetLastErrorDescription());
            }
        }

    }
    public class clsUserFieldsMD  //: IUserFieldsMD, UserFieldsMD
    {
        public clsUserFieldsMD() { }
        
        public SAPbobsCOM.ValidValuesMD ValidValues { get; set; }
    
        public   string Name { get; set; }
         
        public SAPbobsCOM.BoFieldTypes Type { get; set; } 
        public   int Size { get; set; } 
        public   string Description { get; set; }
         
        public SAPbobsCOM.BoFldSubTypes SubType { get; set; } 
        public   string LinkedTable { get; set; } 
        public   string DefaultValue { get; set; } 
        public   string TableName { get; set; } 
        public   int FieldID { get; set; } 
        public   int EditSize { get; set; }
         
        public SAPbobsCOM.BoYesNoEnum Mandatory { get; set; }  
        public   string LinkedUDO { get; set; }

      //  [JsonConverter(typeof(StringEnumConverter))]
      //  public SAPbobsCOM.UDFLinkedSystemObjectTypesEnum LinkedSystemObject { get; set; }
    }
    public class clsUserTablesMD  //: IUserTablesMD, UserTablesMD
    {
        public clsUserTablesMD() { }
       
        public string TableName { get; set; }
        public string TableDescription { get; set; } 
        public SAPbobsCOM.BoUTBTableType TableType { get; set; } 
        public SAPbobsCOM.BoYesNoEnum Archivable { get; set; }
    }
}
