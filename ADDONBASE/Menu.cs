using ADDONBASE.Extensions;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ADDONBASE
{
    class Menu
    {
        SAPbouiCOM.Application SBO_Application;
        //CFLHandler cfl;
        SAPbobsCOM.Company company;


        public Menu(SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company company)
        {

            this.SBO_Application = SBO_Application;
            this.company = company;
            AddMenuItems();

            SBO_Application.MenuEvent += SBO_Application_MenuEvent;
            //      cfl = new CFLHandler(this.SBO_Application);

        }

        internal void AddMenuItems()
        {


            try
            {
                string XmlStr = "";
                var files = System.IO.Directory.GetFiles(@"srf\");
                foreach (var item in files)
                {
                    XmlStr = System.IO.File.ReadAllText(item);

                    try
                    {
                        SBO_Application.LoadBatchActions(System.Xml.Linq.XDocument.Parse(XmlStr).ToString());
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                    try
                    {
                        SAPbouiCOM.MenuItem oMenuItem = SBO_Application.Menus.Item(TC_Menu);

                        //if (!System.IO.File.Exists(Application.StartupPath + "\\workshopRT.png"))
                        //{
                        //    // Set the images for menu folders
                        //    String strPath = System.Windows.Forms.Application.StartupPath + "\\workshop.png";
                        //    Bitmap objBMP = new Bitmap(strPath);
                        //    Size sz = new Size(18, 18);
                        //    objBMP = ResizeImage(objBMP, sz);
                        //    objBMP.Save("workshopRT.gif", System.Drawing.Imaging.ImageFormat.Gif);

                        //}
                        oMenuItem.Image = System.Windows.Forms.Application.StartupPath + "\\icon.png";
                    }
                    catch (Exception Ex) { }
                }

                //       XmlStr = System.IO.File.ReadAllText(@"srf\QLTYMENU.srf");

                // SBO_Application.LoadBatchActions(System.Xml.Linq.XDocument.Parse(XmlStr).ToString());
            }
            catch (Exception er)
            {
            }


        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "1292")
                {
                    var frm = SBO_Application.Forms.ActiveForm;
                    try
                    {
                        if (frm.Items.Item("0_U_G") != null)
                        {
                            var matrix = frm.Items.Item("0_U_G").Specific as SAPbouiCOM.Matrix;
                            matrix.AddRow();


                        }
                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                }

                if (!pVal.BeforeAction && pVal.MenuUID.StartsWith("TC_"))//pVal.MenuUID.StartsWith("u")
                {
                    try
                    {
                        var submenu = this.SBO_Application.Menus.Item("51200").SubMenus;
                        var found = false;
                        var vals = pVal.MenuUID;
                        foreach (SAPbouiCOM.IMenuItem v in submenu)
                        {

                            if (v.String.ToLower().StartsWith(vals.ToLower()))
                            {
                                found = true;
                                v.Activate();
                                break;
                            }

                        }
                        if (!found)
                        {
                            LoadXMLFiles(vals);
                        }

                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                }
                if (!pVal.BeforeAction && _Initializer.UDONames.Contains(pVal.MenuUID))//pVal.MenuUID.StartsWith("u")
                {
                    //47616
                    try
                    {
                        var submenu = this.SBO_Application.Menus.Item("47616").SubMenus;
                        var found = false;
                        var vals = pVal.MenuUID;
                        var valsname = _Initializer.SBO_Application.Menus.Item(pVal.MenuUID).String;
                        foreach (SAPbouiCOM.IMenuItem v in submenu)
                        {
                            if (_Initializer.UDONames.Contains(vals))
                            {

                                if (v.String.Split('-')[0].Trim().Equals(vals) || valsname == v.String)
                                {
                                    found = true;
                                    v.Activate();
                                        break;
                                }

                            }
                        }
                        if (!found)
                        {
                            LoadXMLFiles(vals);
                        }

                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                }
                if (pVal.BeforeAction && pVal.MenuUID.StartsWith("PR_"))
                {
                    var submenu = this.SBO_Application.Menus.Item("51200").SubMenus;

                    var vals = pVal.MenuUID;
                    foreach (SAPbouiCOM.IMenuItem v in submenu)
                    {
                        if (v.String.StartsWith(vals))
                        {
                            v.Activate();
                            break;
                        }
                    }
                }

                if (pVal.BeforeAction && pVal.MenuUID.StartsWith("FRM_"))
                {

                    //var path = System.Windows.Forms.Application.StartupPath + @"\FORMS\" + pVal.MenuUID + ".srf";
                    //var xml = System.IO.File.ReadAllText(path);
                    //SAPbouiCOM.FormCreationParams fcp;
                    //fcp = this.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams) as SAPbouiCOM.FormCreationParams;
                    //fcp.XmlData = System.Xml.Linq.XDocument.Parse(xml).ToString();

                    // var oForm = SBO_Application.Forms.AddEx(fcp);
                    try
                    {
                        _UserFormBase activeForm = (_UserFormBase)System.Reflection.Assembly.GetEntryAssembly().CreateInstance(System.Reflection.Assembly.GetEntryAssembly().GetName().Name + ".BusinessLogic." + pVal.MenuUID);// new FRM_OBSettings();
                        activeForm.Show();
                    }
                    catch (Exception ex)
                    { ex.AppendInLogFile(); }

                }
                if (pVal.BeforeAction && pVal.MenuUID.StartsWith("F_"))
                {
                    var path = System.Windows.Forms.Application.StartupPath + @"\srf\" + pVal.MenuUID + ".srf";
                    var xml = System.IO.File.ReadAllText(path);
                    SAPbouiCOM.FormCreationParams fcp;
                    fcp = this.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams) as SAPbouiCOM.FormCreationParams;
                    fcp.XmlData = System.Xml.Linq.XDocument.Parse(xml).ToString();
                    var oForm = SBO_Application.Forms.AddEx(fcp);
                    oForm.Freeze(true);


                    #region FillGrid
                    {

                        var DT = oForm.DataSources.DataTables.Item("DT_0");
                        DT.Rows.Clear();


                        var recset = this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        recset.DoQuery("select * from  sp_accounts()");
                        recset.MoveFirst();
                        DT.Rows.Add(recset.RecordCount);
                        var ind = 0;
                        while (!recset.EoF)
                        {
                            for (int i = 0; i < DT.Columns.Count; i++)
                            {
                                try
                                {
                                    var value = recset.Fields.Item(DT.Columns.Item(i).Name).Value;
                                    DT.SetValue(DT.Columns.Item(i).Name, ind, value);
                                }
                                catch (Exception ex) { ex.AppendInLogFile(); }
                            }
                            recset.MoveNext();
                            ind++;
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
                        (oForm.Items.Item("Item_5").Specific as SAPbouiCOM.Grid).CollapseLevel = 2;

                        //employee.ExecuteQuery("select code as EmpID,name Employee ,null  Selected from [@PR_OHEM]");
                    }
                    #endregion
                    oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                if (System.Diagnostics.Debugger.IsAttached)
                    SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }
        int FormNum = 0;
        private string LoadXMLFiles(string sFileName)
        {
            var oXmlDoc = default(System.Xml.XmlDocument);
            var oXNode = default(System.Xml.XmlNode);
            var oAttr = default(System.Xml.XmlAttribute);
            string sPath = null;
            string FrmUID = null;
            try
            {
                oXmlDoc = new System.Xml.XmlDocument();

                sPath = Application.StartupPath + "\\FORMS\\" + sFileName + ".srf";

                oXmlDoc.LoadXml(File.ReadAllText(sPath)) ;
                oXNode = oXmlDoc.GetElementsByTagName("form").Item(0);
                oAttr = (System.Xml.XmlAttribute)oXNode.Attributes.GetNamedItem("uid");
                oAttr.Value = oAttr.Value + FormNum;
                FormNum = FormNum + 1;
                _Initializer.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml);
                FrmUID = oAttr.Value;

                return FrmUID;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oXmlDoc = null;
            }
        }

        public static Bitmap ResizeImage(Bitmap imgToResize, Size size)
        {
            try
            {

                Bitmap b = new Bitmap(size.Width, size.Height);
                using (Graphics g = Graphics.FromImage((Image)b))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(imgToResize, 0, 0, size.Width, size.Height);
                }
                return b;
            }
            catch
            {
                Console.WriteLine("Bitmap could not be resized");
                return imgToResize;
            }
        }
        string _TC_Menu = "TC_Menu";
        public string TC_Menu
        {
            get
            {
                return _TC_Menu;
            }
            set
            {
                _TC_Menu = value;
            }
        }
    }
}
