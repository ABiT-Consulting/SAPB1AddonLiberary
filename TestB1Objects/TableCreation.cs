
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Runtime.InteropServices;
using TestB1Objects;

namespace Quality
{
	public class TableCreation
	{
		private long v_RetVal;

		private long v_ErrCode;

		private string v_ErrMsg;

		public string addonName;

		public DateTime startdate;

		public DateTime enddate;

		private Recordset rs;

		private string Qry;
        public void ComponentRoutingMaster()
        {
            try
            {
                this.CRMHeader();
                this.CRMDetails();

                bool flag = !this.UDOExists("uAC_OCRM");
                if (flag)
                {
                    string[,] array = new string[4, 2];
                    array[0, 0] = "DocNum";
                    array[0, 1] = "Document ID";
                    array[1, 0] = "U_docdt";
                    array[1, 1] = "Document Date";
                    array[2, 0] = "U_fgno";
                    array[2, 1] = "FG No.";
                    array[3, 0] = "U_fgname";
                    array[3, 1] = "FG Name";
                    object obj = array;
                    string[,] findField = array;
                    this.RegisterUDO("uAC_OCRM", "ComponentRoutingMaster", BoUDOObjType.boud_Document, findField, "AC_OCRM", "AC_CRM1", "", "", "", BoYesNoEnum.tNO);

                }
            }
            catch (Exception expr_ED)
            {
                ProjectData.SetProjectError(expr_ED);
                Exception ex = expr_ED;
                System.Windows.Forms.MessageBox.Show(ex.Message);
                ProjectData.ClearProjectError();
            }
        }

        public void CRMHeader()
        {
            try
            {
                this.CreateTable("AC_OCRM", "CRM Header", BoUTBTableType.bott_Document);
                this.CreateUserFields("@AC_OCRM", "docdt", "Document Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OCRM", "fgno", "FG No.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OCRM", "sname", "Series Name", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OCRM", "scode", "Series Code", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OCRM", "fgname", "FG Name", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OCRM", "active", "Active/Inactive", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OCRM", "totmc", "Total MachineCost", BoFieldTypes.db_Float, 20L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_OCRM", "totlab", "Total LabourCost", BoFieldTypes.db_Float, 20L, BoFldSubTypes.st_Price, "");
                this.CreateUserFields("@AC_OCRM", "tottool", "Total ToolCost", BoFieldTypes.db_Float, 20L, BoFldSubTypes.st_Price, "");
                this.CreateUserFields("@AC_OCRM", "totcons", "Total ConsumablesCost", BoFieldTypes.db_Float, 20L, BoFldSubTypes.st_Price, "");
                this.CreateUserFields("@AC_OCRM", "totsub", "Total SubcontractingCost", BoFieldTypes.db_Float, 20L, BoFldSubTypes.st_Price, "");
                this.CreateUserFields("@AC_OCRM", "total", "Grand Total", BoFieldTypes.db_Float, 20L, BoFldSubTypes.st_Price, "");

              
            }
            catch (Exception expr_1CB)
            {
                ProjectData.SetProjectError(expr_1CB);
                Exception ex = expr_1CB;
                System.Windows.Forms.MessageBox.Show(ex.Message);
                ProjectData.ClearProjectError();
            }
        }

        public void CRMDetails()
        {
            try
            {
                this.CreateTable("AC_CRM1", "CRM Detail", BoUTBTableType.bott_DocumentLines);
                this.CreateUserFields("@AC_CRM1", "locid", "Location ID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_CRM1", "operid", "Operation ID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_CRM1", "itemid", "Item ID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_CRM1", "SFG", "SFG Flag", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_CRM1", "spec", "Specification", BoFieldTypes.db_Alpha, 250L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_CRM1", "type", "Type", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_CRM1", "toolid", "Tool ID", BoFieldTypes.db_Float, 10L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "seq", "Sequence", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_CRM1", "amccost", "A/C Machine Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "cmccost", "Cumulative Machine Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "alabcost", "A/C Labour Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "clabcost", "Cumulative Labour Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "atoolcost", "A/C Tool Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "ctoolcost", "Cumulative Tool Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "acumcost", "A/C Consumable Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "ccumcost", "Cumulative Consumable Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "subcost", "SubContracting Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "csubcost", "Cumulative SubContracting Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "oprcost", "Operation Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_CRM1", "cumcost", "Cumulative Cost", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");

             
            }
            catch (Exception expr_2F6)
            {
                ProjectData.SetProjectError(expr_2F6);
                Exception ex = expr_2F6;
                System.Windows.Forms.MessageBox.Show(ex.Message);
                ProjectData.ClearProjectError();
            }
        }

        public void OperationMaster()
        {
            try
            {
                this.OperationHeader();
                this.OperationMachine();
                this.OperationConsumablesMachine();
                this.OperationConsumablesTool();
                this.OperationTools();
                bool flag = !this.UDOExists("uAC_OPER");
                if (flag)
                {
                    string[,] array = new string[2, 2];
                    array[0, 0] = "U_operid";
                    array[0, 1] = "Operation ID";
                    array[1, 0] = "U_opername";
                    array[1, 1] = "Operation Name";
                    string[,] findField = array;
                    this.RegisterUDO("uAC_OPER", "OperationMaster", BoUDOObjType.boud_MasterData, findField, "AC_OPER", "AC_OPER1", "AC_OPER2", "AC_OPER3", "AC_OPER4", BoYesNoEnum.tNO);
                }
            }
            catch (Exception expr_B6)
            {
                ProjectData.SetProjectError(expr_B6);
                Exception ex = expr_B6;
                Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                ProjectData.ClearProjectError();
            }
        }
        public void OperationHeader()
        {
            try
            {
                this.CreateTable("AC_OPER", "Operation Master Header", BoUTBTableType.bott_MasterData);
                this.CreateUserFields("@AC_OPER", "operid", "Operation ID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER", "opername", "Operation Name", BoFieldTypes.db_Alpha, 50L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER", "qcreq", "Q/C Required", BoFieldTypes.db_Alpha, 50L, BoFldSubTypes.st_None, "");

            }
            catch (Exception expr_82)
            {
                ProjectData.SetProjectError(expr_82);
                Exception ex = expr_82;
                Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                ProjectData.ClearProjectError();
            }
        }
        
        public void OperationMachine()
        {
            try
            {
                this.CreateTable("AC_OPER1", "Operation Master Machine", BoUTBTableType.bott_MasterDataLines);
                this.CreateUserFields("@QC_AC_OPER1", "mid", "Machine ID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@QC_AC_OPER1", "mname", "Machine Name", BoFieldTypes.db_Alpha, 50L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@QC_AC_OPER1", "consrate", "Consumables Rate", BoFieldTypes.db_Float, 80L, BoFldSubTypes.st_Price, "");

            }
            catch (Exception expr_83)
            {
                ProjectData.SetProjectError(expr_83);
                Exception ex = expr_83;
                Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                ProjectData.ClearProjectError();
            }
        }

        public void OperationConsumablesMachine()
        {
            try
            {
                this.CreateTable("AC_OPER3", "Operation Master Consumables", BoUTBTableType.bott_MasterDataLines);
                this.CreateUserFields("@AC_OPER3", "LineID", "LineID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER3", "DetailID", "DetailID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER3", "itemid", "Itemid", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER3", "unit", "Unit", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER3", "type", "Type", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER3", "qty", "Quantity", BoFieldTypes.db_Float, 10L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_OPER3", "price", "Price", BoFieldTypes.db_Float, 10L, BoFldSubTypes.st_Price, "");
                this.CreateUserFields("@AC_OPER3", "dftprice", "Use Dft Price", BoFieldTypes.db_Alpha, 1L, BoFldSubTypes.st_None, "");

            }
            catch (Exception expr_137)
            {
                ProjectData.SetProjectError(expr_137);
                Exception ex = expr_137;
                Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                ProjectData.ClearProjectError();
            }
        }

        public void OperationConsumablesTool()
        {
            try
            {
                this.CreateTable("AC_OPER4", "Operation Master Consumables 2", BoUTBTableType.bott_MasterDataLines);
                this.CreateUserFields("@AC_OPER4", "LineID", "LineID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER4", "DetailID", "DetailID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER4", "itemid", "Item ID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER4", "unit", "Unit", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER4", "type", "Type", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER4", "qty", "qty", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Quantity, "");
                this.CreateUserFields("@AC_OPER4", "price", "Price", BoFieldTypes.db_Float, 25L, BoFldSubTypes.st_Price, "");


            }
            catch (Exception expr_114)
            {
                ProjectData.SetProjectError(expr_114);
                Exception ex = expr_114;
                Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                ProjectData.ClearProjectError();
            }
        }

        public void OperationTools()
        {
            try
            {
                this.CreateTable("AC_OPER2", "Operation Master Tools", BoUTBTableType.bott_MasterDataLines);
                this.CreateUserFields("@AC_OPER2", "tno", "Tool ID", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_OPER2", "tname", "Tool Name", BoFieldTypes.db_Alpha, 50L, BoFldSubTypes.st_None, "");

            }
            catch (Exception expr_5E)
            {
                ProjectData.SetProjectError(expr_5E);
                Exception ex = expr_5E;
                Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                ProjectData.ClearProjectError();
            }
        }

		public void  TableCreation1()
		{
            //if (System.Configuration.ConfigurationManager.AppSettings["CreateUDFs"] == "true")
            {
                this.v_ErrMsg = "";
                this.addonName = "Quality";
                this.ParameterMaster();
                this.createMethodRef();
                this.LayoutParameterMaster();
                this.ControlPlan();
                this.InwardInspection();
                //this.InitialInspection();
                this.InprocessInspection();
                this.FinalInspection();
                //this.PreDispatchinspection();
                this.ItemMaster();
                this.OperationMaster();
                this.ComponentRoutingMaster();
            }
		}

		

		public void ParameterMaster()
		{
			try
			{
				this.createparamtable();
				this.createparamdetail();
				bool flag = !this.UDOExists("uAC_PM");
				if (flag)
				{
					string[,] array = new string[1, 2];
					array[0, 0] = "U_pname1";
					array[0, 1] = "Parameter Name";
					string[,] findField = array;
					this.RegisterUDO("uAC_PM", "ParameterMaster", BoUDOObjType.boud_Document, findField, "AC_PrmMtr", "AC_PrmMtr1", "", "", BoYesNoEnum.tNO);
				}
			}
			catch
			{
			}
		}
        public void createMethodRef()
        {
            try
            {
                this.CreateTable("AC_Methodref", "Method Referance Master", BoUTBTableType.bott_MasterData);
                this.CreateUserFields("@AC_Methodref", "MDesc", "Master Data Desc", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
                bool flag = !this.UDOExists("uAC_Methodref");
                if (flag)
                {
                    string[,] array = new string[1, 2];
                    array[0, 0] = "U_MDesc";
                   // array[0, 1] = "Parameter Name";
                    string[,] findField = array;
                    this.RegisterUDO("uAC_Methodref", "MethodReferance", BoUDOObjType.boud_MasterData, findField, "AC_Methodref", "", "", "", BoYesNoEnum.tNO);
                }
            }
            catch (Exception expr_73)
            {
            }

        }
		public void createparamtable()
		{
			this.CreateTable("AC_PrmMtr", "Quality Parameter Master", BoUTBTableType.bott_Document);
			this.CreateUserFields("@AC_PrmMtr", "ptype", "Param Type", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_PrmMtr", "pname1", "Parameter Name", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_PrmMtr", "subtype", "Sub Type", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
		}

		public void createparamdetail()
		{
			this.CreateTable("AC_PrmMtr1", "Quality Parameter Detail", BoUTBTableType.bott_DocumentLines);
			this.CreateUserFields("@AC_PrmMtr1", "specify1", "specifications", BoFieldTypes.db_Memo, 800L, BoFldSubTypes.st_None, "");
		}

		public void LayoutParameterMaster()
		{
			try
			{
				this.createlayoutbasic();
				this.createlayoutdetail();
				bool flag = !this.UDOExists("uAC_QLTYPLN");
				if (flag)
				{
					string[,] array = new string[2, 2];
					array[0, 0] = "DocNum";
					array[0, 1] = "DocNum";
					array[1, 0] = "U_itemid";
					array[1, 1] = "Item Id";
					string[,] findField = array;
					this.RegisterUDO("uAC_QLTYPLN", "LayoutParameterMaster", BoUDOObjType.boud_Document, findField, "AC_QLTYPLN", "AC_QLTYPLN1", "", "", BoYesNoEnum.tNO);
				}
			}
			catch (Exception expr_8D)
			{
			}
		}

		public void createlayoutbasic()
		{
			try
			{
				this.CreateTable("AC_QLTYPLN", "Quality Layout Master Header", BoUTBTableType.bott_Document);
				this.CreateUserFields("@AC_QLTYPLN", "docdt", "Doc.Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "itemid", "Item Id", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "itemdesc", "Item Description", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "revno", "Rev.No", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "revdt", "Rev.Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "preby", "Prepared By", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "prebycode", "Prepared By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "custreq", "Customer Req.", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "oapp", "Other Approvals", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN", "oappdt", "Other Approval Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_153)
			{
			}
		}

		public void createlayoutdetail()
		{
			try
			{
				this.CreateTable("AC_QLTYPLN1", "Quality Layout Master Detail", BoUTBTableType.bott_DocumentLines);
                this.CreateUserFields("@AC_QLTYPLN1", "ptype", "Param Type.", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_QLTYPLN1", "pname1", "Param Name", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN1", "spec1", "Specification", BoFieldTypes.db_Memo, 800L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_QLTYPLN1", "MDesc", "Description", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
                this.CreateUserFields("@AC_QLTYPLN1", "accpcre", "Accepted Criteria", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN1", "insp", "Inspection Method", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN1", "rem", "Remarks", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN1", "instrumentno", "Instrument No", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN1", "specchars", "Spec.Char", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN1", "inspfreq", "Inspection Frequency", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN1", "bysupplier", "By Supplier", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_QLTYPLN1", "byus", "By US", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_15C)
			{
			}
		}

		public void ControlPlan()
		{
			try
			{
				this.createcpbasic();
				this.createcpdetail();
				bool flag = !this.UDOExists("uAC_CP");
				if (flag)
				{
					string[,] array = new string[2, 2];
					array[0, 0] = "DocNum";
					array[0, 1] = "DocNum";
					array[1, 0] = "U_itemid";
					array[1, 1] = "Item Id";
					string[,] findField = array;
					this.RegisterUDO("uAC_CP", "ControlPlan", BoUDOObjType.boud_Document, findField, "AC_CP", "AC_CP1", "", "", BoYesNoEnum.tNO);
				}
			}
			catch (Exception expr_8D)
			{
			}
		}

		public void createcpbasic()
		{
			try
			{
				this.CreateTable("AC_CP", "Quality Control Plan Header", BoUTBTableType.bott_Document);
				this.CreateUserFields("@AC_CP", "ptype", "Plan Type", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "itemid", "Item Id", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "itemdesc", "Item Description", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "crmno", "CRM No.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "revno", "Rev.No", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "revdt", "Rev.Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "preby", "Prepared By", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "prebycode", "Prepared By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "samplesize", "Sample Size", BoFieldTypes.db_Numeric, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "frequency", "Frequency", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "freqval", "Frequency Value", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "custreq", "Customer Requirement", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "oapp", "Other Approval", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "cqapp", "Cust.Quality Approval", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "cqappdt", "Cust.Quality App.Dt", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "ceapp", "Cust.Engg.Approval", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "ceappdt", "Cust.Engg.App.Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "srpapp", "Supplier/Plant Approval", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "srpappdt", "Supplier/Plant App.Dt", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "docdt", "Doc.Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "oappdt", "Other App.Dt", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "cteam", "Core Team", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "phone", "Phone", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "cperson", "Contact Person", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "carno", "Cust.Ass.Ref.No", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "aarno", "Our Ass.Ref.No", BoFieldTypes.db_Alpha, 50L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "suppid", "Supplier Id", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "partyid", "Party Id", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP", "operid", "Stage of Plan.", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_3AF)
			{
			}
		}

		public void createcpdetail()
		{
			try
			{
				this.CreateTable("AC_CP1", "Quality Control Plan Detail", BoUTBTableType.bott_DocumentLines);
				this.CreateUserFields("@AC_CP1", "ptype", "Param Type", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "pname1", "Param Name", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "spec1", "Specification", BoFieldTypes.db_Memo, 800L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "cmethod", "Control Method", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "imethod", "Inspection Method", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "ftype", "Freq. Type", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "freq", "Freq. Value", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "smpsize", "Sample Size", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "mcno", "Machine No", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "toolno", "Tool No", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "rplan", "Reaction Plan", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "corract", "Corractive Action", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_CP1", "specchars", "Special Characteristics", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_1BC)
			{
			}
		}

		public void InwardInspection()
		{
			try
			{
				this.createinwardbasic();
				this.createinwarddetail();
				this.createinwardsgrid();
				this.createInwardBatchSerilTrack();
				bool flag = !this.UDOExists("uAC_INWARD");
				if (flag)
				{
					string[,] array = new string[2, 2];
					array[0, 0] = "DocNum";
					array[0, 1] = "DocNum";
					array[1, 0] = "U_suppid";
					array[1, 1] = "Supplier Id";
					string[,] findField = array;
					this.RegisterUDO("uAC_INWARD", "InwardInspection", BoUDOObjType.boud_Document, findField, "AC_INWRD", "AC_INWRD1", "AC_INWRD2", "AC_INWRD3", BoYesNoEnum.tNO);
				}
			}
			catch (Exception expr_9B)
			{
			}
		}

		public void createinwardbasic()
		{
			try
			{
				this.CreateTable("AC_INWRD", "Quality Inward Header", BoUTBTableType.bott_Document);
				this.CreateUserFields("@AC_INWRD", "itype", "Inward Type", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "suppid", "SupplierID", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "suppname", "SupplierName", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "grnno", "GRN No.", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "grndt", "GRN Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "docdt", "Inspection Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "rewloc", "Rework Location", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "rejloc", "Rejected Location", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "Pgrnno", "POSTGRN No.", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "samloc", "Sample Location", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "irrec", "Ins.Report Received", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "irdt", "Ins.Report Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "tcrec", "Test Certificate Rec.", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "tcdt", "Test Certificate Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "corractrep", "Corrective Action Report.", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "corractrepdt", "Corrective Action Report Date.", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "rem", "Remarks", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "dact", "Disposal Action", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "chkby", "Checked By", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "chkbycode", "Checked By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "pendreason", "Reason For Pending", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "authby", "Authorized by", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "authby1", "Auth Code", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "inspby", "Inspect By", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "inspby1", "Inspect Code", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "InspNo", "Inspection Number", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "TaxNo", "Tax Certificate Number", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "Approval", "Approval", BoFieldTypes.db_Alpha, 1L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INWRD", "QCNUMS", "QC Number", BoFieldTypes.db_Alpha, 50L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_3B6)
			{
			}
		}

		public void createinwarddetail()
		{
			this.CreateTable("AC_INWRD1", "Quality Inward Detail", BoUTBTableType.bott_DocumentLines);
			this.CreateUserFields("@AC_INWRD1", "itemid", "Item ID", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "itemdesc", "Item Desc", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "unit", "Unit", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "recqty", "Rec. Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_INWRD1", "insqty", "Inspec.Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_INWRD1", "accpqty", "Accepted Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_INWRD1", "accpwdev", "Deviation Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_INWRD1", "rewqty", "Rework Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_INWRD1", "rejqty", "Rejected Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_INWRD1", "rate", "Rate", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Rate, "");
			this.CreateUserFields("@AC_INWRD1", "reason", "Rejection Reason", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "recloc", "Received Location", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "samploc", "Sample Location", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "acceploc", "Accepted Location", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "Analy", "Analysis No", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "status", "Status", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "status1", "Status1", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD1", "SamQty", "Sample Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
		}

		public void createinwardsgrid()
		{
			this.CreateTable("AC_INWRD2", "Quality Inward SubGrid", BoUTBTableType.bott_DocumentLines);
			this.CreateUserFields("@AC_INWRD2", "basicid", "Basic ID", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD2", "detailid", "Detail Id", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD2", "pname1", "Parameter Name", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD2", "spec1", "Specifications", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD2", "observation", "observations", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD2", "remarks", "Remarks", BoFieldTypes.db_Alpha, 120L, BoFldSubTypes.st_None, "");
		}

		public void createInwardBatchSerilTrack()
		{
			this.CreateTable("AC_INWRD3", "Quality InWard Serial Number", BoUTBTableType.bott_DocumentLines);
			this.CreateUserFields("@AC_INWRD3", "chkbox", "Check Box", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "Type", "Type", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "itemid", "Item Code", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "serial", "Serial No.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "batchqty", "Batch Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_INWRD3", "location", "Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "sysserial", "System Serial", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "suppserial", "SuppSerial", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "class", "Batch or Serial Class", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "fromloc", "From Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_INWRD3", "Quantity", "Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_INWRD3", "WhsType", "WhsType", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
		}

		public void InitialInspection()
		{
			try
			{
				this.createinitialbasic();
				this.createinitialdetail();
				bool flag = !this.UDOExists("QC_INITIAL");
				if (flag)
				{
					string[,] array = new string[2, 2];
					array[0, 0] = "DocNum";
					array[0, 1] = "Document No";
					array[1, 0] = "U_itemid";
					array[1, 1] = "Item Id";
					string[,] findField = array;
					this.RegisterUDO("QC_INITIAL", "Initial Inspection", BoUDOObjType.boud_Document, findField, "QC_INITIALBASIC", "QC_INITIALDETAIL", "", "", BoYesNoEnum.tNO);
				}
			}
			catch (Exception expr_8D)
			{
			}
		}

		public void createinitialbasic()
		{
			try
			{
				this.CreateTable("QC_INITIALBASIC", "Quality Initial Header", BoUTBTableType.bott_Document);
				this.CreateUserFields("@QC_INITIALBASIC", "shift", "Shift", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "inspdt", "Inspection Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "itemid", "Item Id.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "itemdesc", "Item Desc", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "shopno", "Shop Order No", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "shopdt", "Shop Order Dt.", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "preby", "Prepared By", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "prebycode", "Prepared By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "oper", "Operator Name", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "opercode", "Operator Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "mcno", "Machine No", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "comm", "Comments", BoFieldTypes.db_Alpha, 200L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "predt", "Prepared Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "recom", "Recom. For Prod.", BoFieldTypes.db_Alpha, 200L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "opname", "Oper.Name.", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "status", "Status", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "ncremarks", "NC Remarks.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALBASIC", "prodqty", "Produced Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@QC_INITIALBASIC", "accpqty", "Accpted Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@QC_INITIALBASIC", "rejqty", "Rejected Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@QC_INITIALBASIC", "SamQty", "Sample Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			}
			catch (Exception expr_2BC)
			{
			}
		}

		public void createinitialdetail()
		{
			try
			{
				this.CreateTable("QC_INITIALDETAIL", "Quality Initial Detail", BoUTBTableType.bott_DocumentLines);
				this.CreateUserFields("@QC_INITIALDETAIL", "ptype", "Param Type.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "pchk", "Param's to be chked.", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "spec1", "Specification", BoFieldTypes.db_Alpha, 60L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "samp1", "Samp1.", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "samp2", "Samp2.", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "samp3", "Samp3.", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "samp4", "Samp4.", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "samp5", "Samp5.", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "time1", "Time1", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "opdt", "opdt", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "rem", "Remarks.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "time2", "Time2", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "time3", "Time3", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "time4", "Time4", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_INITIALDETAIL", "time5", "Time5", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_1F8)
			{
				ProjectData.SetProjectError(expr_1F8);
				Exception ex = expr_1F8;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void InprocessInspection()
		{
			try
			{
				this.createinprocessbasic();
				this.createinprocessdetail();
				bool flag = !this.UDOExists("uAC_INPROCS");
				if (flag)
				{
					string[,] array = new string[2, 2];
					array[0, 0] = "DocNum";
					array[0, 1] = "Document No";
					array[1, 0] = "U_itemid";
					array[1, 1] = "Item Id";
					string[,] findField = array;
					this.RegisterUDO("uAC_INPROCS", "InprocessInspection", BoUDOObjType.boud_Document, findField, "AC_INPROCS", "AC_INPROCS1", "", "", BoYesNoEnum.tNO);
				}
			}
			catch (Exception expr_8D)
			{
				ProjectData.SetProjectError(expr_8D);
				Exception ex = expr_8D;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createinprocessbasic()
		{
			try
			{
				this.CreateTable("AC_INPROCS", "Quality Inprocess Header", BoUTBTableType.bott_Document);
				this.CreateUserFields("@AC_INPROCS", "shift", "Shift", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "inspdt", "Inspection Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "itemid", "Item Id.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "itemdesc", "Item Desc", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "shopno", "Shop Order No", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "shopdt", "Shop Order Dt.", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "preby", "Prepared By", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "prebycode", "Prepared By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "predt", "Prepared Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "crmno", "Process Sheet No", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "ncdet", "N.C.Det.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "action", "Action.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "prodqty", "Prod.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_INPROCS", "accpqty", "Accp.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_INPROCS", "rewqty", "Rew.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_INPROCS", "rejqty", "Rej.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_INPROCS", "inspby", "Insp.By", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "inspbycode", "Insp.By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "chkby", "Checked By", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "chkbycode", "Checked By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "insptime", "Inspection time.", BoFieldTypes.db_Date, 10L, BoFldSubTypes.st_Time, "");
				this.CreateUserFields("@AC_INPROCS", "sampleqty", "Sample Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_INPROCS", "opname", "Operation Name", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "mcno", "Machine No", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "oper", "Operator Name", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "opercode", "Operator Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "sname", "Supplier Name", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "BatchN", "Batch Number", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "operid", "Stage of Plan", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS", "opernm", "Stage of Plan name", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_3DA)
			{
				ProjectData.SetProjectError(expr_3DA);
				Exception ex = expr_3DA;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createinprocessdetail()
		{
			try
			{
				this.CreateTable("AC_INPROCS1", "Quality Inprocess Detail", BoUTBTableType.bott_DocumentLines);
				this.CreateUserFields("@AC_INPROCS1", "ptype", "Param Type.", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "pname1", "Param.Name.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "spec1", "Specification", BoFieldTypes.db_Memo, 800L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "t1read", "t1read", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "t2read", "t2read", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "t3read", "t3read", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "t4read", "t4read", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "t5read", "t5read", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "time1", "Time1", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "time2", "Time2", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "time3", "Time3", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "time4", "Time4", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_INPROCS1", "time5", "Time5", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_1BC)
			{
				ProjectData.SetProjectError(expr_1BC);
				Exception ex = expr_1BC;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void GetFinanacial_year()
		{
			this.rs = (Recordset)Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
			this.Qry = "select top 1 OACP.\"FinancYear\" from OACP order by \"AbsEntry\" desc";
			this.rs.DoQuery(this.Qry);
			this.startdate = Conversions.ToDate(this.rs.Fields.Item("FinancYear").Value);
			this.enddate = this.startdate.AddYears(1);
		}

		public void FinalInspection()
		{
			try
			{
				this.createfinalbasic();
				this.createfinaldetail();
				this.createFinalBatchSerilTrack();
				bool flag = !this.UDOExists("uAC_FINAL");
				if (flag)
				{
					string[,] array = new string[2, 2];
					array[0, 0] = "DocNum";
					array[0, 1] = "Document No";
					array[1, 0] = "U_itemid";
					array[1, 1] = "Item Id";
					string[,] findField = array;
					this.RegisterUDO("uAC_FINAL", "FinalInspection", BoUDOObjType.boud_Document, findField, "AC_FINAL", "AC_FINAL1", "AC_FINAL2", "", BoYesNoEnum.tNO);
				}
			}
			catch (Exception expr_94)
			{
				ProjectData.SetProjectError(expr_94);
				Exception ex = expr_94;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createfinalbasic()
		{
			try
			{
				this.CreateTable("AC_FINAL", "Quality Fianl Header", BoUTBTableType.bott_Document);
				this.CreateUserFields("@AC_FINAL", "dept", "Department", BoFieldTypes.db_Alpha, 25L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "inspdt", "Inspection Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "itemid", "Item Id.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "itemdesc", "Item Desc", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "shopno", "Shop Order No", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "shopnm", "Shop Order Num", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "shopdt", "Shop Order Dt.", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "crmno", "Process Sheet No", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "cpno", "Control Plan No", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "spins", "Special Ins.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "rem", "Remarks", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "sampqty", "Samp.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_FINAL", "inspby", "Insp.By", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "inspbycode", "Insp.By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "status", "Status", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "prodqty", "Prod.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_FINAL", "accpqty", "Accp.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_FINAL", "rewqty", "Rew.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_FINAL", "rejqty", "Rej.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@AC_FINAL", "fromloc", "From Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "accploc", "Accpted Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "rejloc", "Rejected Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "rewloc", "Rework Location ", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "samploc", "Sample Location ", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL", "Approval", "Approval", BoFieldTypes.db_Alpha, 1L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_339)
			{
				ProjectData.SetProjectError(expr_339);
				Exception ex = expr_339;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createfinaldetail()
		{
			try
			{
				this.CreateTable("AC_FINAL1", "Quality Fianl Detail", BoUTBTableType.bott_DocumentLines);
				this.CreateUserFields("@AC_FINAL1", "ptype", "Param Type.", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "pname1", "Param.Name.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "spec1", "Specification", BoFieldTypes.db_Memo, 800L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "mmin", "Min.Meas.", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "mmax", "Max Meas.", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m1", "m1", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m2", "m2", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m3", "m3", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m4", "m4", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m5", "m5", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m6", "m6", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m7", "m7", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m8", "m8", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m9", "m9", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@AC_FINAL1", "m10", "m10", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_1FC)
			{
				ProjectData.SetProjectError(expr_1FC);
				Exception ex = expr_1FC;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createFinalBatchSerilTrack()
		{
			this.CreateTable("AC_FINAL2", "Quality Final Serial Number", BoUTBTableType.bott_DocumentLines);
			this.CreateUserFields("@AC_FINAL2", "chkbox", "Check Box", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "Type", "Type", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "itemid", "Item Code", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "serial", "Serial No.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "batchqty", "Batch Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_FINAL2", "location", "Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "sysserial", "System Serial", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "suppserial", "SuppSerial", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "class", "Batch or Serial Class", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "fromloc", "From Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@AC_FINAL2", "Quantity", "Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@AC_FINAL2", "WhsType", "WhsType", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
		}

		public void PreDispatchinspection()
		{
			try
			{
				this.createpdbasic();
				this.createpddetail();
				this.createtesting();
				this.createDisptchBatchSerilTrack();
				bool flag = !this.UDOExists("QC_PREDISPATCH");
				if (flag)
				{
					string[,] array = new string[2, 2];
					array[0, 0] = "DocNum";
					array[0, 1] = "DocNum";
					array[1, 0] = "U_itemid";
					array[1, 1] = "Item Id";
					string[,] findField = array;
					this.RegisterUDO("QC_PREDISPATCH", "Quarantine Inspection", BoUDOObjType.boud_Document, findField, "QC_PDBASIC", "QC_PDDETAIL", "QC_TESTING", "QC_DSPSERIAL", BoYesNoEnum.tNO);
				}
			}
			catch (Exception expr_9B)
			{
				ProjectData.SetProjectError(expr_9B);
				Exception ex = expr_9B;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createpdbasic()
		{
			try
			{
				this.CreateTable("QC_PDBASIC", "Quality Predispatch Header", BoUTBTableType.bott_Document);
				this.CreateUserFields("@QC_PDBASIC", "reportdt", "Report Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "itemid", "Item Id.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "itemdesc", "Item Desc", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "shopno", "Shop Order No", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "shopdt", "Shop Order Dt.", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "prodqty", "Prod.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@QC_PDBASIC", "accpqty", "Accp.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@QC_PDBASIC", "rewqty", "Rew.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@QC_PDBASIC", "rejqty", "Rej.Qty.", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@QC_PDBASIC", "fromloc", "From Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "accploc", "Accpted Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "rejloc", "Rejected Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "rewloc", "Rework Location ", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "samploc", "Sample Location ", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "partyid", "Party Id", BoFieldTypes.db_Alpha, 15L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "partydesc", "Party Desc.", BoFieldTypes.db_Alpha, 100L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "samqty", "Sample Qty", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
				this.CreateUserFields("@QC_PDBASIC", "authby", "Authorized By", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "authbycode", "Authorized By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "authdt", "Authorized Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "inspby", "Inspected By", BoFieldTypes.db_Alpha, 80L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "inspdt", "Inspected Date", BoFieldTypes.db_Date, 0L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "inspbycode", "Inspected By Code", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "remarks", "Cust.Remarks.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDBASIC", "Approval", "Approval", BoFieldTypes.db_Alpha, 1L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_334)
			{
				ProjectData.SetProjectError(expr_334);
				Exception ex = expr_334;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createpddetail()
		{
			try
			{
				this.CreateTable("QC_PDDETAIL", "Quality Predispatch Detail", BoUTBTableType.bott_DocumentLines);
				this.CreateUserFields("@QC_PDDETAIL", "pname1", "Param.Name.", BoFieldTypes.db_Alpha, 254L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "spec1", "Specification", BoFieldTypes.db_Memo, 800L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs1", "Obs1", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs2", "Obs2", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs3", "Obs3", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs4", "Obs4", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs5", "Obs5", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs6", "Obs6", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs7", "Obs7", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs8", "Obs8", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs9", "Obs9", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "obs10", "Obs10", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_PDDETAIL", "custobs", "Cust.Obs", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_1BC)
			{
				ProjectData.SetProjectError(expr_1BC);
				Exception ex = expr_1BC;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createtesting()
		{
			try
			{
				this.CreateTable("QC_TESTING", "Quality Other Testing", BoUTBTableType.bott_DocumentLines);
				this.CreateUserFields("@QC_TESTING", "tname", "Test Name.", BoFieldTypes.db_Alpha, 40L, BoFldSubTypes.st_None, "");
				this.CreateUserFields("@QC_TESTING", "status", "Status", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			}
			catch (Exception expr_56)
			{
				ProjectData.SetProjectError(expr_56);
				Exception ex = expr_56;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
		}

		public void createDisptchBatchSerilTrack()
		{
			this.CreateTable("QC_DSPSERIAL", "Quality PreDisptch SNo.", BoUTBTableType.bott_DocumentLines);
			this.CreateUserFields("@QC_DSPSERIAL", "chkbox", "Check Box", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "Type", "Type", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "itemid", "Item Code", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "serial", "Serial No.", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "batchqty", "Batch Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@QC_DSPSERIAL", "location", "Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "sysserial", "System Serial", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "suppserial", "SuppSerial", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "class", "Batch or Serial Class", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "fromloc", "From Location", BoFieldTypes.db_Alpha, 20L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("@QC_DSPSERIAL", "Quantity", "Quantity", BoFieldTypes.db_Float, 0L, BoFldSubTypes.st_Quantity, "");
			this.CreateUserFields("@QC_DSPSERIAL", "WhsType", "WhsType", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
		}

		public void ItemMaster()
		{
			this.CreateUserFieldsComboBox("OITM", "qcinsp", "QC Inspection Yes Or No", BoFieldTypes.db_Alpha, 10L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("ODRF", "IKEY", "InternalKey", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
			this.CreateUserFields("IBT1", "bsdPrNo", "Base Pr No", BoFieldTypes.db_Alpha, 30L, BoFldSubTypes.st_None, "");
		}

		public bool UDOExists(string code)
		{
			GC.Collect();
			UserObjectsMD userObjectsMD = (UserObjectsMD)Program.oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
			bool byKey = userObjectsMD.GetByKey(code);
			Marshal.ReleaseComObject(userObjectsMD);
			return byKey;
		}

		public bool CreateTable(string TableName, string TableDesc, BoUTBTableType TableType)
		{
			bool result = false;
			string text = "";
			try
			{
				bool flag = !this.TableExists(TableName);
				if (flag)
				{
					Program.oapp.StatusBar.SetText("Creating Table " + TableName + " ...................", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
					UserTablesMD userTablesMD = (UserTablesMD)Program.oCompany.GetBusinessObject(BoObjectTypes.oUserTables);
					userTablesMD.TableName = TableName;
					userTablesMD.TableDescription = TableDesc;
					userTablesMD.TableType = TableType;
					long num = (long)userTablesMD.Add();
					flag = (num != 0L);
					if (flag)
					{
						SAPbobsCOM.ICompany arg_94_0 = Program.oCompany;
						long num3=0;
						int num2 = checked((int)num3);
						arg_94_0.GetLastError(out num2, out text);
						num3 = (long)num2;
						Program.oapp.StatusBar.SetText(string.Concat(new string[]
						{
							"Failed to Create Table ",
							TableDesc,
							Conversions.ToString(num3),
							" ",
							text
						}), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
						Marshal.ReleaseComObject(userTablesMD);
						result = false;
					}
					else
					{
						Program.oapp.StatusBar.SetText(string.Concat(new string[]
						{
							"[",
							TableName,
							"] - ",
							TableDesc,
							" Created Successfully!!!"
						}), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
						Marshal.ReleaseComObject(userTablesMD);
						result = true;
					}
				}
				else
				{
					GC.Collect();
					result = false;
				}
			}
			catch (Exception expr_16A)
			{
				ProjectData.SetProjectError(expr_16A);
				Exception ex = expr_16A;
				Program.oapp.StatusBar.SetText(string.Concat(new string[]
				{
					this.addonName,
					":> ",
					ex.Message,
					" @ ",
					ex.Source
				}), BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
			return result;
		}

		public bool ColumnExists(string TableName, string FieldID)
		{
			bool result=true;
			try
			{
				Recordset recordset = (Recordset)Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
				bool flag = true;
				recordset.DoQuery(string.Concat(new string[]
				{
					"Select 1 from \"CUFD\" Where \"TableID\"='",
					Strings.Trim(TableName),
					"' and \"AliasID\"='",
					Strings.Trim(FieldID),
					"'"
				}));
				bool eoF = recordset.EoF;
				if (eoF)
				{
					flag = false;
				}
				Marshal.ReleaseComObject(recordset);
				GC.Collect();
				result = flag;
			}
			catch (Exception expr_86)
			{
				ProjectData.SetProjectError(expr_86);
				Exception ex = expr_86;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
			return result;
		}

		public bool UDFExists(string TableName, string FieldID)
		{
			bool result=true;
			try
			{
				Recordset recordset = (Recordset)Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
				bool flag = true;
				recordset.DoQuery(string.Concat(new string[]
				{
                    "Select 1 from \"CUFD\" Where \"TableID\"='",
					Strings.Trim(TableName),
                    "' and \"AliasID\"='",
					Strings.Trim(FieldID),
					"'"
				}));
				bool eoF = recordset.EoF;
				if (eoF)
				{
					flag = false;
				}
				Marshal.ReleaseComObject(recordset);
				GC.Collect();
				result = flag;
			}
			catch (Exception expr_86)
			{
				ProjectData.SetProjectError(expr_86);
				Exception ex = expr_86;
				Program.oapp.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
				ProjectData.ClearProjectError();
			}
			return result;
		}

		public bool TableExists(string TableName)
		{
			UserTablesMD userTablesMD = (UserTablesMD)Program.oCompany.GetBusinessObject(BoObjectTypes.oUserTables);
			bool byKey = userTablesMD.GetByKey(TableName);
			Marshal.ReleaseComObject(userTablesMD);
			return byKey;
		}

		public bool CreateUserFieldsComboBox(string TableName, string FieldName, string FieldDescription, BoFieldTypes type, long size = 0L, BoFldSubTypes subType = BoFldSubTypes.st_None, string LinkedTable = "")
		{
			bool result=true;
			try
			{
				bool flag = !TableName.StartsWith("@");
				if (flag)
				{
					bool flag2 = !this.UDFExists(TableName, FieldName);
					if (flag2)
					{
						UserFieldsMD userFieldsMD = (UserFieldsMD)Program.oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
						try
						{
							userFieldsMD.TableName = TableName;
							userFieldsMD.Name = FieldName;
							userFieldsMD.Description = FieldDescription;
							userFieldsMD.Type = type;
							flag2 = (type != BoFieldTypes.db_Date);
							if (flag2)
							{
								flag = (size != 0L);
								if (flag)
								{
									userFieldsMD.Size = checked((int)size);
								}
							}
							flag2 = (subType != BoFldSubTypes.st_None);
							if (flag2)
							{
								userFieldsMD.SubType = subType;
							}
							userFieldsMD.ValidValues.Value = "N";
							userFieldsMD.ValidValues.Description = "N";
							userFieldsMD.ValidValues.Add();
							userFieldsMD.ValidValues.Value = "Y";
							userFieldsMD.ValidValues.Description = "Y";
							userFieldsMD.ValidValues.Add();
							userFieldsMD.DefaultValue = "N";
							flag2 = (Operators.CompareString(LinkedTable, "", false) != 0);
							if (flag2)
							{
								userFieldsMD.LinkedTable = LinkedTable;
							}
							this.v_RetVal = (long)userFieldsMD.Add();
							flag2 = (this.v_RetVal != 0L);
							if (flag2)
							{
								SAPbobsCOM.ICompany arg_166_0 = Program.oCompany;
								int num = checked((int)this.v_ErrCode);
								arg_166_0.GetLastError(out num, out this.v_ErrMsg);
								this.v_ErrCode = (long)num;
								Program.oapp.StatusBar.SetText(string.Concat(new string[]
								{
								"Failed to add UserField ",
								FieldDescription,
								" - ",
								Conversions.ToString(this.v_ErrCode),
								" ",
								this.v_ErrMsg
								}), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
								Marshal.ReleaseComObject(userFieldsMD);
								result = false;
							}
							else
							{
								Program.oapp.StatusBar.SetText(" & TableName & - " + FieldDescription + " added successfully!!!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
								Marshal.ReleaseComObject(userFieldsMD);
								result = true;
							}
						}finally
						{
							if(userFieldsMD != null)
							{
                                Marshal.ReleaseComObject(userFieldsMD);
                            }
						}
					}
					else
					{
						result = false;
					}
				}
			}
			catch (Exception expr_21F)
			{
				ProjectData.SetProjectError(expr_21F);
				Exception ex = expr_21F;
				Program.oapp.MessageBox(ex.Message, 1, "Ok", "", "");
				ProjectData.ClearProjectError();
			}
			return result;
		}

		public bool CreateUserFields(string TableName, string FieldName, string FieldDescription, BoFieldTypes type, long size = 0L, BoFldSubTypes subType = BoFldSubTypes.st_None, string LinkedTable = "")
		{
			bool result=true;
			try
			{
				bool flag = TableName.StartsWith("@");
				if (flag)
				{
					bool flag2 = !this.ColumnExists(TableName, FieldName);
					if (flag2)
					{
						UserFieldsMD userFieldsMD = (UserFieldsMD)Program.oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
						try
						{
							userFieldsMD.TableName = TableName;
							userFieldsMD.Name = FieldName;
							userFieldsMD.Description = FieldDescription;
							userFieldsMD.Type = type;
							flag2 = (type != BoFieldTypes.db_Date);
							if (flag2)
							{
								flag = (size != 0L);
								if (flag)
								{
									userFieldsMD.Size = checked((int)size);
								}
							}
							flag2 = (subType != BoFldSubTypes.st_None);
							if (flag2)
							{
								userFieldsMD.SubType = subType;
							}
							flag2 = (Operators.CompareString(LinkedTable, "", false) != 0);
							if (flag2)
							{
								userFieldsMD.LinkedTable = LinkedTable;
							}
							this.v_RetVal = (long)userFieldsMD.Add();
							flag2 = (this.v_RetVal != 0L);
							if (flag2)
							{
								SAPbobsCOM.ICompany arg_F9_0 = Program.oCompany;
								int num = checked((int)this.v_ErrCode);
								arg_F9_0.GetLastError(out num, out this.v_ErrMsg);
								this.v_ErrCode = (long)num;
								Program.oapp.StatusBar.SetText("Failed to add UserField masterid" + Conversions.ToString(this.v_ErrCode) + " " + this.v_ErrMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
								Marshal.ReleaseComObject(userFieldsMD);
								result = false;
							}
							else
							{
								Program.oapp.StatusBar.SetText(string.Concat(new string[]
								{
								"[",
								TableName,
								"] - ",
								FieldDescription,
								" added successfully!!!"
								}), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
								Marshal.ReleaseComObject(userFieldsMD);
								result = true;
							}
						}finally
						{
							if(userFieldsMD != null)
							{
                                Marshal.ReleaseComObject(userFieldsMD);
                            }
						}
					}
					else
					{
						result = false;
					}
				}
				else
				{
					bool flag2 = !TableName.StartsWith("@");
					if (flag2)
					{
						flag = !this.UDFExists(TableName, FieldName);
						if (flag)
						{
							UserFieldsMD userFieldsMD2 = (UserFieldsMD)Program.oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
							try
							{
								userFieldsMD2.TableName = TableName;
								userFieldsMD2.Name = FieldName;
								userFieldsMD2.Description = FieldDescription;
								userFieldsMD2.Type = type;
								flag2 = (type != BoFieldTypes.db_Date);
								if (flag2)
								{
									flag = (size != 0L);
									if (flag)
									{
										userFieldsMD2.Size = checked((int)size);
									}
								}
								flag2 = (subType != BoFldSubTypes.st_None);
								if (flag2)
								{
									userFieldsMD2.SubType = subType;
								}
								flag2 = (Operators.CompareString(LinkedTable, "", false) != 0);
								if (flag2)
								{
									userFieldsMD2.LinkedTable = LinkedTable;
								}
								this.v_RetVal = (long)userFieldsMD2.Add();
								flag2 = (this.v_RetVal != 0L);
								if (flag2)
								{
									SAPbobsCOM.ICompany arg_2B2_0 = Program.oCompany;
									int num = checked((int)this.v_ErrCode);
									arg_2B2_0.GetLastError(out num, out this.v_ErrMsg);
									this.v_ErrCode = (long)num;
									Program.oapp.StatusBar.SetText(string.Concat(new string[]
									{
									"Failed to add UserField ",
									FieldDescription,
									" - ",
									Conversions.ToString(this.v_ErrCode),
									" ",
									this.v_ErrMsg
									}), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
									Marshal.ReleaseComObject(userFieldsMD2);
									result = false;
								}
								else
								{
									Program.oapp.StatusBar.SetText(" & TableName & - " + FieldDescription + " added successfully!!!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
									Marshal.ReleaseComObject(userFieldsMD2);
									result = true;
								}
							}
							finally
							{
								if(userFieldsMD2 != null)
								{
                                    Marshal.ReleaseComObject(userFieldsMD2);
                                }
							}
						}
						else
						{
							result = false;
						}
					}
				}
			}
			catch (Exception expr_36C)
			{
				ProjectData.SetProjectError(expr_36C);
				Exception ex = expr_36C;
				Program.oapp.MessageBox(ex.Message, 1, "Ok", "", "");
				ProjectData.ClearProjectError();
			}
			return result;
		}

		public bool RegisterUDO(string UDOCode, string UDOName, BoUDOObjType UDOType, string[,] FindField, string UDOHTableName, string UDODTableName = "", string ChildTable = "", string ChildTable1 = "", BoYesNoEnum LogOption = BoYesNoEnum.tNO)
		{
			bool flag = false;
			bool result=true;
			try
			{
				result = false;
				UserObjectsMD userObjectsMD = (UserObjectsMD)Program.oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
				userObjectsMD.CanCancel = BoYesNoEnum.tYES;
				userObjectsMD.CanClose = BoYesNoEnum.tYES;
				userObjectsMD.CanDelete = BoYesNoEnum.tYES;
				userObjectsMD.CanFind = BoYesNoEnum.tYES;
				userObjectsMD.CanLog = BoYesNoEnum.tNO;
				userObjectsMD.CanYearTransfer = BoYesNoEnum.tYES;
				userObjectsMD.ManageSeries = BoYesNoEnum.tYES;
				userObjectsMD.Code = UDOCode;
				userObjectsMD.Name = UDOName;
				userObjectsMD.TableName = UDOHTableName;

                userObjectsMD.RebuildEnhancedForm = BoYesNoEnum.tNO;
                userObjectsMD.EnableEnhancedForm = BoYesNoEnum.tYES;
                userObjectsMD.CanCreateDefaultForm = BoYesNoEnum.tYES;
				bool flag2 = Operators.CompareString(UDODTableName, "", false) != 0;
				if (flag2)
				{
					userObjectsMD.ChildTables.TableName = UDODTableName;
					userObjectsMD.ChildTables.Add();
				}
				flag2 = (Operators.CompareString(ChildTable, "", false) != 0);
				if (flag2)
				{
					userObjectsMD.ChildTables.TableName = ChildTable;
					userObjectsMD.ChildTables.Add();
				}
				flag2 = (Operators.CompareString(ChildTable1, "", false) != 0);
				if (flag2)
				{
					userObjectsMD.ChildTables.TableName = ChildTable1;
					userObjectsMD.ChildTables.Add();
				}
				flag2 = (LogOption == BoYesNoEnum.tYES);
				if (flag2)
				{
					userObjectsMD.LogTableName = "A" + UDOHTableName;
				}
				userObjectsMD.ObjectType = UDOType;
				short arg_145_0 = 0;
				short num = checked((short)(FindField.GetLength(0) - 1));
				short num2 = arg_145_0;
				while (true)
				{
					short arg_195_0 = num2;
					short num3 = num;
					if (arg_195_0 > num3)
					{
						break;
					}
					flag2 = (num2 > 0);
					if (flag2)
					{
						userObjectsMD.FindColumns.Add();
					}
					userObjectsMD.FindColumns.ColumnAlias = FindField[(int)num2, 0];
					userObjectsMD.FindColumns.ColumnDescription = FindField[(int)num2, 1];
					num2 += 1;
				}
				flag2 = (userObjectsMD.Add() == 0);
				if (flag2)
				{
					result = true;
					flag2 = Program.oCompany.InTransaction;
					if (flag2)
					{
						Program.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
					}
                    try
                    {
                        var key = Program.oCompany.GetNewObjectKey();
                        var Value = "Y";
                        string query = string.Format("update  OUDO set \"CanDefForm\"='{1}' where \"Code\" = '{0}'", key, Value);
                        var oRecset = Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        oRecset .DoQuery(query);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecset);
                        oRecset = null;
                        GC.Collect();
                    }
                    catch
                    { }
					Program.oapp.StatusBar.SetText(string.Concat(new string[]
					{
						"Successfully Registered UDO >",
						UDOCode,
						">",
						UDOName,
						" >",
						Program.oCompany.GetLastErrorDescription()
					}), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				}
				else
				{
					Program.oapp.StatusBar.SetText(string.Concat(new string[]
					{
						"Failed to Register UDO >",
						UDOCode,
						">",
						UDOName,
						" >",
						Program.oCompany.GetLastErrorDescription()
					}), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
					result = false;
				}
				Marshal.ReleaseComObject(userObjectsMD);
				GC.Collect();
				flag2 = (!flag & Program.oCompany.InTransaction);
				if (flag2)
				{
					Program.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
				}
			}
			catch (Exception expr_2B0)
			{
				ProjectData.SetProjectError(expr_2B0);
				bool flag2 = Program.oCompany.InTransaction;
				if (flag2)
				{
					Program.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
				}
				ProjectData.ClearProjectError();
			}
			return result;
		}

        public bool RegisterUDO(string UDOCode, string UDOName, BoUDOObjType UDOType, string[,] FindField, string UDOHTableName, string UDODTableName = "", string ChildTable = "", string ChildTable1 = "", string ChildTable2 = "", BoYesNoEnum LogOption = BoYesNoEnum.tNO)
        {
            bool flag = false;
            bool result = true;
            try
            {

                result = false;
                UserObjectsMD userObjectsMD = (UserObjectsMD)Program.oCompany.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                userObjectsMD.CanCancel = BoYesNoEnum.tYES;
                userObjectsMD.CanClose = BoYesNoEnum.tYES;
                userObjectsMD.CanCreateDefaultForm = BoYesNoEnum.tNO;
                userObjectsMD.CanDelete = BoYesNoEnum.tYES;
                userObjectsMD.CanFind = BoYesNoEnum.tYES;
                userObjectsMD.CanLog = BoYesNoEnum.tNO;
                userObjectsMD.CanYearTransfer = BoYesNoEnum.tYES;
                userObjectsMD.ManageSeries = BoYesNoEnum.tYES;
                userObjectsMD.Code = UDOCode;
                userObjectsMD.Name = UDOName;
                userObjectsMD.TableName = UDOHTableName;
                bool flag2 = Operators.CompareString(UDODTableName, "", false) != 0;
                if (flag2)
                {
                    userObjectsMD.ChildTables.TableName = UDODTableName;
                    userObjectsMD.ChildTables.Add();
                }
                flag2 = (Operators.CompareString(ChildTable, "", false) != 0);
                if (flag2)
                {
                    userObjectsMD.ChildTables.TableName = ChildTable;
                    userObjectsMD.ChildTables.Add();
                }
                flag2 = (Operators.CompareString(ChildTable1, "", false) != 0);
                if (flag2)
                {
                    userObjectsMD.ChildTables.TableName = ChildTable1;
                    userObjectsMD.ChildTables.Add();
                }
                flag2 = (Operators.CompareString(ChildTable2, "", false) != 0);
                if (flag2)
                {
                    userObjectsMD.ChildTables.TableName = ChildTable2;
                    userObjectsMD.ChildTables.Add();
                }
                flag2 = (LogOption == BoYesNoEnum.tYES);
                if (flag2)
                {
                    userObjectsMD.LogTableName = "A" + UDOHTableName;
                }
                userObjectsMD.ObjectType = UDOType;
                short arg_145_0 = 0;
                short num = checked((short)(FindField.GetLength(0) - 1));
                short num2 = arg_145_0;
                while (true)
                {
                    short arg_195_0 = num2;
                    short num3 = num;
                    if (arg_195_0 > num3)
                    {
                        break;
                    }
                    flag2 = (num2 > 0);
                    if (flag2)
                    {
                        userObjectsMD.FindColumns.Add();
                    }
                    userObjectsMD.FindColumns.ColumnAlias = FindField[(int)num2, 0];
                    userObjectsMD.FindColumns.ColumnDescription = FindField[(int)num2, 1];
                    num2 += 1;
                }
                flag2 = (userObjectsMD.Add() == 0);
                if (flag2)
                {
                    result = true;
                    flag2 = Program.oCompany.InTransaction;
                    if (flag2)
                    {
                        Program.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                    }
                    Program.oapp.StatusBar.SetText(string.Concat(new string[]
                    {
                        "Successfully Registered UDO >",
                        UDOCode,
                        ">",
                        UDOName,
                        " >",
                        Program.oCompany.GetLastErrorDescription()
                    }), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    Program.oapp.StatusBar.SetText(string.Concat(new string[]
                    {
                        "Failed to Register UDO >",
                        UDOCode,
                        ">",
                        UDOName,
                        " >",
                        Program.oCompany.GetLastErrorDescription()
                    }), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    result = false;
                }
                Marshal.ReleaseComObject(userObjectsMD);
                GC.Collect();
                flag2 = (!flag & Program.oCompany.InTransaction);
                if (flag2)
                {
                    Program.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
            }
            catch (Exception expr_2B0)
            {
                ProjectData.SetProjectError(expr_2B0);
                bool flag2 = Program.oCompany.InTransaction;
                if (flag2)
                {
                    Program.oCompany.EndTransaction(BoWfTransOpt.wf_RollBack);
                }
                ProjectData.ClearProjectError();
            }
            return result;
        }
  
    }
}
