using Microsoft.VisualBasic.Logging;
using Quality;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
namespace TestB1Objects
{
    public partial class TestForm : Form
    {
        SAPbouiCOM.Application oapp;
        SAPbobsCOM.Company company;
        string ids = "";
        public TestForm()
        {
            InitializeComponent();
        }

        private void TestForm_Load(object sender, EventArgs e)
        {
            oapp = Program.oapp;
            company = oapp.Company.GetDICompany() as SAPbobsCOM.Company;
            this.textBox1.Text = string.Format("Server = {0}{5}License Server= {1}{5}CompanyDB= {2}{5}DbUserName= {3}{5}UserName= {4} ", company.Server, company.LicenseServer, company.CompanyDB, company.DbUserName, company.UserName, Environment.NewLine);
            this.tbServer.Text = company.Server;
            this.tbLicenseServer.Text = company.LicenseServer;
            this.tbSLDServer.Text = company.SLDServer ;
            this.tbCompanyDB.Text = company.CompanyDB;
            this.tbDBUser.Text = company.DbUserName;
            this.tbUserName.Text = company.UserName;
            this.tbServerType.Text = company.DbServerType.ToString();
            

        }


        private void btnGetXML_Click(object sender, EventArgs e)
        {
            company.XmlExportType = (SAPbobsCOM.BoXmlExportTypes)Enum.Parse(typeof(SAPbobsCOM.BoXmlExportTypes), cbxmltype.SelectedItem.ToString());

            var obj = company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Convert.ToInt32(tbObjID.Text)) as dynamic;
            obj.GetByKey(Convert.ToInt32(tbKey.Text));
            tbXML.Text = XDocument.Parse(obj.GetAsXML()).ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            company.XmlExportType = (SAPbobsCOM.BoXmlExportTypes)Enum.Parse(typeof(SAPbobsCOM.BoXmlExportTypes), cbxmltype.SelectedItem.ToString());

            var obj = company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Convert.ToInt32(tbObjID.Text)) as dynamic;
            if (obj is SAPbobsCOM.Documents)
            {
                var doc = (SAPbobsCOM.Documents)obj;
                doc.GetByKey(Convert.ToInt32(tbKey.Text));
                if (doc.Cancel() != 0)
                {
                    if (company.GetLastErrorCode() == -5006)
                    {
                        var doc1 = doc.CreateCancellationDocument();
                        if (doc1.Add() == 0)
                        {
                            MessageBox.Show("Successfully added");
                            return;
                        }
                    }
                    MessageBox.Show(company.GetLastErrorCode() + " " + company.GetLastErrorDescription());
                }
                else
                {

                    MessageBox.Show("Successfully added");
                }
            }

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                var path = Path.GetTempFileName();
                var dir = Directory.GetDirectoryRoot(path);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                System.IO.File.WriteAllText(path, tbXML.Text);
                var obj = company.GetBusinessObjectFromXML(path, 0) as dynamic;
                var i = obj.Add();
                if (i == 0)
                {
                    MessageBox.Show("Successfully added at " + company.GetNewObjectKey());
                    Console.WriteLine("Successfully added at " + company.GetNewObjectKey());
                }
                else
                {
                    Console.WriteLine(company.GetLastErrorCode() + " " + company.GetLastErrorDescription());
                    MessageBox.Show(company.GetLastErrorCode() + " " + company.GetLastErrorDescription());

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            var CompanyNew = new SAPbobsCOM.Company();
            CompanyNew.Server = tbServer.Text;
            CompanyNew.LicenseServer = this.tbLicenseServer.Text;
            CompanyNew.SLDServer = this.tbSLDServer.Text;
            CompanyNew.CompanyDB = this.tbCompanyDB.Text;
            CompanyNew.DbUserName = this.tbDBUser.Text;
            CompanyNew.DbPassword = tbSystemPassword.Text;
            CompanyNew.UserName = this.tbUserName.Text;
            CompanyNew.Password = tbManagerPassword.Text;
            CompanyNew.DbServerType = company.DbServerType;
            var i = CompanyNew.Connect();
            if (i == 0)
            {
                MessageBox.Show("Company Connected.");

            }
            else
            {
                MessageBox.Show(CompanyNew.GetLastErrorCode() + " " + CompanyNew.GetLastErrorDescription());
            }

        }








        private void btnUDFs_Click(object sender, EventArgs e)
        {

        }
        public bool updateAP(string str, char pro, SAPbobsCOM.Company oCompany)
        {
            bool res = true;
            var query2 = String.Format("update OWTM set Active = '{1}' WHERE WtmCode in ({0})", str, pro);
            var rec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            try { rec.DoQuery(query2); } catch (Exception ex) { res = false; } finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rec); }

            return res;
        }
        private void BtnDisableApprovals_Click(object sender, EventArgs e)
        {
            var sesionId = "admin";
            var query = String.Format("Select T0.WtmCode from OWTM T0 INNER JOIN WTM1 T1 ON T0.WtmCode=T1.WtmCode INNER JOIN WTM3 T2 ON T1.WtmCode=T2.WtmCode INNER JOIN OUSR T3 ON  T1.UserID=T3.USERID AND T3.USER_CODE<>'{0}' WHERE T2.TransType='1470000113' AND T0.Active='Y' and T0.USER_CODE <> 'manager'", sesionId);
            var rec = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            try
            {
                rec.DoQuery(query);
                while (!rec.EoF) { ids = ids + "," + Convert.ToString(rec.Fields.Item("WtmCode").Value); rec.MoveNext(); }
                if (!string.IsNullOrEmpty(ids)) { ids = ids.Substring(1); }
            }
            catch (Exception ex) { }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(rec); }
            updateAP(ids, 'N', company);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            updateAP(ids, 'Y', company);
            ids = "";

        }

        private void Clear_Click(object sender, EventArgs e)
        {
            tbXML.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;


            var xml = company.GetBusinessObjectXmlSchema((SAPbobsCOM.BoObjectTypes)Convert.ToInt32(tbObjID.Text));
            tbXML.Text = XDocument.Parse(xml).ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(tbXML.Text);
        }

        private void BtnItem_Click(object sender, EventArgs e)
        {
            var recset = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(@"select ""ItemCode"" from oitm where ""ItemCode"" like '100-%' and ""ItmsGrpCod"" <>'100'");
            var itemmaster = company.GetBusinessObject(BoObjectTypes.oItems) as SAPbobsCOM.Items;

            while (!recset.EoF)
            {
                itemmaster.GetByKey(recset.Fields.Item(0).Value.ToString().Trim());
                itemmaster.ItemsGroupCode = 100;
                var i = itemmaster.Update();
                if (i == 0)
                {
                    Console.WriteLine("Successfully updated Item " + recset.Fields.Item(0).Value.ToString().Trim());
                }
                else
                {
                    Console.WriteLine(company.GetLastErrorDescription());
                }
                recset.MoveNext();
            }
            Marshal.ReleaseComObject(recset);
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            var query = @"SELECT  [ItemCode]
      ,[ItemName]
      ,[ItmsGrpCod]
      ,[UoMGroupEntry]
      ,[PrchseItem]
      ,[SellItem]
      ,[InvntItem]
      ,[posted]
  FROM [staging].[dbo].[oitm] where isnull(posted ,0)= 0";
            var recset = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            var recset1 = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(query);
            while (!recset.EoF)
            {
                var ItemCode = recset.Fields.Item("ItemCode").Value.ToString();
                var ItemName = recset.Fields.Item("ItemName").Value.ToString();
                var ItmsGrpCod = Convert.ToInt32(recset.Fields.Item("ItmsGrpCod").Value);
                var UoMGroupEntry = Convert.ToInt32(recset.Fields.Item("UoMGroupEntry").Value);
                var PrchseItem = recset.Fields.Item("PrchseItem").Value.ToString();
                var SellItem = recset.Fields.Item("SellItem").Value.ToString();
                var InvntItem = recset.Fields.Item("InvntItem").Value.ToString();
                var oitm = company.GetBusinessObject(BoObjectTypes.oItems) as SAPbobsCOM.Items;
                oitm.ItemCode = ItemCode;
                oitm.ItemName = ItemName;
                oitm.ItemsGroupCode = ItmsGrpCod;
                oitm.UoMGroupEntry = UoMGroupEntry;
                if (PrchseItem.ToUpper() == "Y")
                    oitm.PurchaseItem = BoYesNoEnum.tYES;
                else
                    oitm.PurchaseItem = BoYesNoEnum.tNO;

                if (SellItem.ToUpper() == "Y")
                    oitm.SalesItem = BoYesNoEnum.tYES;
                else
                    oitm.SalesItem = BoYesNoEnum.tNO;

                if (InvntItem.ToUpper() == "Y")
                    oitm.InventoryItem = BoYesNoEnum.tYES;
                else
                    oitm.InventoryItem = BoYesNoEnum.tNO;
                if (oitm.Add() == 0)
                {
                    try
                    {
                        Log log = new Log();
                        log.WriteEntry("Added ItemCode " + ItemCode);
                        recset1.DoQuery(string.Format(" update [staging].[dbo].[oitm] set posted = '1' where ItemCode = '{0}'", ItemCode));
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        Log log = new Log();
                        log.WriteEntry("Failed ItemCode " + ItemCode + " " + company.GetLastErrorDescription());
                        recset1.DoQuery(string.Format(" update [staging].[dbo].[oitm] set posted = '2' where ItemCode = '{0}'", ItemCode));
                    }
                    catch { }

                    MessageBox.Show(company.GetLastErrorDescription());
                }
                Marshal.ReleaseComObject(oitm);
                recset.MoveNext();
            }
            Marshal.ReleaseComObject(recset);
            Marshal.ReleaseComObject(recset1);
        }

        private void btnUDF_Click(object sender, EventArgs e)
        {
            TableCreation tblcreation = new TableCreation();
            var added = tblcreation.CreateUserFields(tbTableName.Text, tbFieldName.Text, tbDescription.Text, (BoFieldTypes)Enum.Parse(typeof(BoFieldTypes), cbType.Text), Convert.ToInt64(tbSize.Text), (BoFldSubTypes)Enum.Parse(typeof(BoFldSubTypes), cbSubType.Text), tbLinkTable.Text);
            if (added)
            { MessageBox.Show("Successfully Added"); }
            else
                MessageBox.Show("Failed to Add");
        }

        private void BtnBarcode_Click(object sender, EventArgs e)
        {
            var query = @"Select ""ItemCode"", ""DocEntry"" from oitm where ""frozenFor"" = 'N'";
            var recset = company.GetBusinessObject(BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            var items = company.GetBusinessObject(BoObjectTypes.oItems) as SAPbobsCOM.Items;
            recset.DoQuery(query);
            while (!recset.EoF)
            {
                items.GetByKey(recset.Fields.Item("ItemCode").Value.ToString());
                items.ForeignName = recset.Fields.Item("ItemCode").Value.ToString();

                items.BarCode = "01-" + recset.Fields.Item("DocEntry").Value.ToString();
                items.Update();
                recset.MoveNext();
            }
            Marshal.ReleaseComObject(recset);
            GC.Collect();
            MessageBox.Show("Completed Successfully");
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }
        public bool Order2Invoice(int DocEntry, out string message)
        {
            var result = true; message = "success";
            var order = company.GetBusinessObject(BoObjectTypes.oOrders) as SAPbobsCOM.Documents;
            var invoice = company.GetBusinessObject(BoObjectTypes.oInvoices) as SAPbobsCOM.Documents;
            try
            {
                var st = order.GetByKey(DocEntry);
                if (st)
                {
                    company.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;
                    var str = order.GetAsXML();
                    XDocument doc = XDocument.Parse(str);
                    #region Modify Xml


                    var header = doc.Descendants("Documents").FirstOrDefault().Descendants("row").FirstOrDefault();
                    header.Descendants("ReserveInvoice").FirstOrDefault().Value = "tYES";
                    header.Descendants("DocObjectCode").FirstOrDefault().Value = "13";
                    header.Descendants("NumberOfInstallments").Remove();
                    header.Descendants("Series").Remove();

                    foreach (var row in doc.Descendants("Document_Lines").FirstOrDefault().Descendants("row"))
                    {
                        if (row.Descendants("BaseEntry").Count() > 0) { row.Descendants("BaseEntry").First().Value = doc.Descendants("DocEntry").First().Value; } else { row.Add(new XElement("BaseEntry", doc.Descendants("DocEntry").First().Value)); }
                        if (row.Descendants("BaseLine").Count() > 0) { row.Descendants("BaseLine").First().Value = row.Descendants("LineNum").FirstOrDefault().Value; } else { row.Add(new XElement("BaseLine", row.Descendants("LineNum").FirstOrDefault().Value)); }
                        if (row.Descendants("BaseType").Count() > 0) { row.Descendants("BaseType").First().Value = "17"; } else { row.Add(new XElement("BaseType", "17")); }

                    }

                    doc.Descendants("AdmInfo").FirstOrDefault().Descendants("Object").FirstOrDefault().Value = "13";
                    doc.Descendants("DocEntry").Remove();
                    doc.Descendants("DocNum").Remove();
                    doc.Descendants().Where(x => x.Value == "").Remove();

                    #endregion
                    tbXML.Text = doc.ToString();
                    #region creating invoice from xml
                    var path = Path.GetTempFileName();
                    var dir = Directory.GetDirectoryRoot(path);
                    if (!Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }
                    System.IO.File.WriteAllText(path, doc.ToString());
                    var obj = company.GetBusinessObjectFromXML(path, 0) as dynamic;
                    var i = obj.Add();
                    if (i != 0)
                    {
                        message = company.GetLastErrorDescription();
                        result = false;
                    }
                    #endregion
                }
                else
                {
                    result = false;
                    message = "Doc not Found";
                }
            }
            catch (Exception ex)
            {
                result = false;
                message = ex.Message;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(order);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(invoice);
            }
            return result;
        }


        const string Connection = "Connection";
        const string _Server = "Server";
        const string _CompanyDB = "CompanyDB";
        const string _DbUserName = "DbUserName";
        const string _DbPassword = "DbPassword";
        const string _UserName = "UserName";
        const string _Password = "Password";
        const string _LicenseServer = "LicenseServer";
        const string _language = "language";
        const string _DbServerType = "DbServerType";
        const string _SLDServer = "SLDServer";
        const string _EncFileName = "cnx.p";
        const string psw = "!@#908!112aFgKts";
        private void btncnxp_Click(object sender, EventArgs e)
        {
            try
            {
                #region connect Company
                var CompanyNew = new SAPbobsCOM.Company();
                CompanyNew.Server = tbServer.Text;
                CompanyNew.LicenseServer = this.tbLicenseServer.Text;
                CompanyNew.CompanyDB = this.tbCompanyDB.Text;
                CompanyNew.DbUserName = this.tbDBUser.Text;
                CompanyNew.DbPassword = tbSystemPassword.Text;
                CompanyNew.UserName = this.tbUserName.Text;
                CompanyNew.Password = tbManagerPassword.Text;
                CompanyNew.DbServerType = company.DbServerType;
                var i = CompanyNew.Connect();
                if (i == 0)
                {
                    MessageBox.Show("Company Connected.");

                }
                else
                {
                    MessageBox.Show(CompanyNew.GetLastErrorCode() + " " + CompanyNew.GetLastErrorDescription());
                }
                #endregion
            }
            catch { }
            var xml = "<Connection> <Server>SAMIR</Server> <CompanyDB>SQL_CFM</CompanyDB> <DbUserName>sa</DbUserName><DbPassword>P@ssw0rd</DbPassword><UserName>manager</UserName>  <Password>manager</Password>  <LicenseServer>Samir:30000</LicenseServer>  <language>3</language><SLDServer>3</SLDServer>  <DbServerType>7</DbServerType></Connection>";

            XDocument doc = XDocument.Parse(xml);
            var connection = doc.Element(Connection);
            connection.Element(_Server).Value = tbServer.Text;
            connection.Element(_CompanyDB).Value = this.tbCompanyDB.Text;
            connection.Element(_DbUserName).Value = this.tbDBUser.Text;
            connection.Element(_DbPassword).Value = tbSystemPassword.Text;
            connection.Element(_UserName).Value = this.tbUserName.Text;
            connection.Element(_Password).Value = tbManagerPassword.Text;
            connection.Element(_LicenseServer).Value = this.tbLicenseServer.Text;
            // connection.Element(_language).Value = CompanyNew.language.ToString();
            connection.Element(_DbServerType).Value = company.DbServerType.ToString();
            //connection.Element(_SLDServer).Value = this.Company.SLDServer.ToString();
            // password = tbSystemPassword.Text;
            var txt = connection.ToString();
            txt.EncryptMeToFile(psw, "cnx.p");
        }
        /// <summary>
        /// Add using DI Server
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            var xml = tbXML.Text;
            int err = -1;
            string err_msg = "";
            string key = "";
            MessageBox.Show(err_msg + " at " + key);
        }

        private void tbUpdate_Click(object sender, EventArgs e)
        {
            var temp = Path.GetTempFileName();
            System.IO.File.WriteAllText(temp, tbXML.Text);
            company.XmlExportType = (SAPbobsCOM.BoXmlExportTypes)Enum.Parse(typeof(SAPbobsCOM.BoXmlExportTypes), cbxmltype.SelectedItem.ToString());

            dynamic obj = company.GetBusinessObjectFromXML(temp, 0);
            var i = obj.Update();
            if (i != 0)
            {
                Debug.Write(company.GetLastErrorDescription());
                MessageBox.Show(company.GetLastErrorDescription());
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            //Show message box if the user really wants to delete udos
            if (MessageBox.Show("Are you sure you want to delete the UDOs?", "Delete UDOs", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var d = @"update oadm set ""U_datavers""";

                //query to get oudo
                var query = $@"select ""Code"" from OUDO ";
                //recordset from company
                var recset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                try { recset.DoQuery(d); } catch { }
                recset.DoQuery(query);
                List<string> codes = new List<string>();
                while (!recset.EoF)
                {
                    codes.Add(recset.Fields.Item("Code").Value.ToString());
                    recset.MoveNext();
                }
                Marshal.ReleaseComObject(recset);
                GC.Collect();
                //loop on recordset
                bool success = true;
                foreach (var code in codes)
                {
                    var obj = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
                    //get object by docentry
                    var b = obj.GetByKey(code);
                    //delete object
                    if (b)
                    {
                        int i = obj.Remove();
                        if (i != 0)
                        {
                            success = false;
                            Debug.Write(company.GetLastErrorDescription());
                            MessageBox.Show(company.GetLastErrorDescription());
                        }

                    }
                    Marshal.ReleaseComObject(obj);
                    GC.Collect();
                }
                if (success)
                    MessageBox.Show("UDOs Deleted Successfully");
                else
                    MessageBox.Show("Error in UDO Delete");
            }

        }

        private void btnRemoveUDFs_Click(object sender, EventArgs e)
        {
            //Show message box if the user really wants to delete udf
            if (MessageBox.Show("Are you sure you want to delete the UDFs?", "Delete UDFs", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var tablename = tbTableName.Text;
                var udfname = tbFieldName.Text;
                var query = $@"select ""FieldID"" from CUFD where ""TableID"" = '{tablename}' and ""AliasID"" like '{udfname}%'";
                var recset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                recset.DoQuery(query);
                recset.MoveFirst();
                int udfid = Convert.ToInt32(recset.Fields.Item("FieldID").Value);
                Marshal.ReleaseComObject(recset);
                var obj = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields) as SAPbobsCOM.UserFieldsMD;
                try
                {
                    var b = obj.GetByKey(tablename, udfid);
                    if (b)
                    {
                        int i = obj.Remove();
                        if (i != 0)
                        {
                            Debug.Write(company.GetLastErrorDescription());
                            MessageBox.Show(company.GetLastErrorDescription());
                        }
                        else
                        {
                            MessageBox.Show("UDF removed successfully");
                        }
                    }
                }
                finally
                {
                    if (obj != null)
                        Marshal.ReleaseComObject(obj);
                }
            }
        }

        private void btnDeleteUDFs_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to delete All UDFs?", "Delete UDFs", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var query = $@"select ""TableID"", ""FieldID"" from CUFD where ""AliasID"" not like 'B1SYS%'";
                var recset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                recset.DoQuery(query);
                recset.MoveFirst();
                List<(string tableid, int fieldid)> list = new List<(string tableid, int fieldid)>();
                while(!recset.EoF)
                {
                    list.Add((recset.Fields.Item("TableID").Value.ToString(), Convert.ToInt32(recset.Fields.Item("FieldID").Value)));
                    recset.MoveNext();
                }
                Marshal.ReleaseComObject(recset);
                GC.Collect();
                bool success = true;
                foreach (var row in list)
                {
                    var tablename = row.tableid;
                    var udfid = row.fieldid;
                    var obj = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields) as SAPbobsCOM.UserFieldsMD;
                    try
                    {
                        var b = obj.GetByKey(tablename, udfid);
                        if (b)
                        {
                            int i = obj.Remove();
                            if (i != 0)
                            {
                                success = false;
                                Debug.Write(company.GetLastErrorDescription());
                                MessageBox.Show(company.GetLastErrorDescription());
                            }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(obj);
                        GC.Collect();
                    }
                }
                if (success)
                    MessageBox.Show("UDFs deleted successfully");
                else
                    MessageBox.Show("Error while removing UDFs");

            }

        }

        private void btnDeleteUDTs_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to delete All UDTs?", "Delete UDTs", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var query = $@"select ""TableName"" from OUTB";
                var recset = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                recset.DoQuery(query);
                recset.MoveFirst();
                var codes = new List<string>();
                while (!recset.EoF)
                {
                    codes.Add(recset.Fields.Item("TableName").Value.ToString());
                    recset.MoveNext();
                }
                Marshal.ReleaseComObject(recset);
                GC.Collect();
                bool success = true;
                foreach (var code in codes)
                {
                    var obj = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables) as SAPbobsCOM.UserTablesMD;
                    var b = obj.GetByKey(code);
                    if (b)
                    {
                        int i = obj.Remove();
                        if (i != 0)
                        {
                            success = false;
                            Debug.Write(company.GetLastErrorDescription());
                            MessageBox.Show(company.GetLastErrorDescription());
                        }
                    }
                    Marshal.ReleaseComObject(obj);
                    GC.Collect();
                }
                if (success)
                {
                    MessageBox.Show("UDTs removed successfully");
                }
                else
                {
                    MessageBox.Show("Error while removing UDTs");
                }




            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var form = Program.oapp.Forms.ActiveForm;
            form.Items.Item("Item_6").Click();
            form.Items.Item("strDate").Click();
            Program.oapp.SendKeys("20.5.23");
            Program.oapp.SendKeys("{TAB}");
            form.Items.Item("endDate").Click();
            Program.oapp.SendKeys("t");
            Program.oapp.SendKeys("{TAB}");

            form.Items.Item("btnfilter").Click();
            form.Items.Item("Item_6").Click();


        }

        private void btnCreateUDF_FROM_XML_Click(object sender, EventArgs e)
        {
            var company = Program.oCompany;
            #region Create UDFS
            try
            {
                udcreator createudo = new udcreator(company);

                Program.oapp.SetStatusBarMessage("Process Started");
                var path = "udf.xml";
                createudo.createUDFFromXML(path); 
                Program.oapp.SetStatusBarMessage("Process Completed");
            }
            catch (Exception Ex) { Program.oapp.SetStatusBarMessage(Ex.Message); 
                Program.oapp.SetStatusBarMessage("Process Ended with Error");

            }
            #endregion
        }
    }
}
