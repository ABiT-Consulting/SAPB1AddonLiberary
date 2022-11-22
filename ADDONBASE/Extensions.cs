using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Xml.Linq;
namespace ADDONBASE.Extensions
{
    public static class Extensions
    {

        internal static string getObjectKeyFromXML(String XML)
        {
            try
            {
                XML = XML.Replace("<?xml version=\"1.0\" encoding=\"UTF-16\" ?>", "");
                XML = XML.Replace(" ", "");
                XML = XML.Replace("CenterCode", "DocEntry");
                XML = XML.Replace("JdtNum", "DocEntry");
                XML = XML.Replace("Code", "DocEntry");
                XML = XML.Replace("AbsoluteEntry", "DocEntry");
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            var xdoc = System.Xml.Linq.XDocument.Parse(XML);
            var value = xdoc.Root.Element("DocEntry").Value;
            return value;
        }

        public static void ApplyXMLLanguage(this SAPbouiCOM.IForm CurrentForm)
        {
            try
            {
                string FileName = string.Format("{0}\\Translate\\{1}.xml", Environment.CurrentDirectory, CurrentForm.TypeEx);
                var doc = XDocument.Load(FileName);
                if (doc.Element("FORM").Attribute("OLDCaption").Value != doc.Element("FORM").Attribute("NEWCaption").Value)
                {
                    CurrentForm.Title = doc.Element("FORM").Attribute("NEWCaption").Value;
                }
                Parallel.ForEach(doc.Descendants("ITEM"), (xe) =>
                {
                   //foreach (XElement xe in doc.Descendants("ITEM"))
                   //{
                   if (CurrentForm.Items.Item(xe.Attribute("UID").Value).Specific is SAPbouiCOM.Matrix)
                    {
                        var matrix = CurrentForm.Items.Item(xe.Attribute("UID").Value).Specific as SAPbouiCOM.Matrix;
                        var columns = xe.Descendants("COLUMN");
                        Parallel.ForEach(columns, (item) =>

                        // foreach (XElement item in columns)
                        {
                                    if (item.Attribute("OLDCaption").Value != item.Attribute("NEWCaption").Value)

                                        matrix.Columns.Item(item.Attribute("UID").Value).TitleObject.Caption = item.Attribute("NEWCaption").Value;
                                });
                    }
                    else
                    if (xe.Attribute("OLDCaption").Value != xe.Attribute("NEWCaption").Value)
                    {
                        if (CurrentForm.Items.Item(xe.Attribute("UID").Value).Specific is SAPbouiCOM.StaticText)
                        {

                            (CurrentForm.Items.Item(xe.Attribute("UID").Value).Specific as SAPbouiCOM.StaticText).Caption = xe.Attribute("NEWCaption").Value;
                        }
                        else
                            if (CurrentForm.Items.Item(xe.Attribute("UID").Value).Specific is SAPbouiCOM.Folder)
                        {

                            (CurrentForm.Items.Item(xe.Attribute("UID").Value).Specific as SAPbouiCOM.Folder).Caption = xe.Attribute("NEWCaption").Value;
                        }
                    }

                });
            }
            catch (Exception ex)
            {
                ex.AppendInLogFile();
            }
        }

        public static void MakeLanguageFile(this SAPbouiCOM.IForm CurrentForm)
        {
            Task tsk = new Task(() =>
            {
                if (!Directory.Exists("Language"))
                {
                    Directory.CreateDirectory("Language");
                }
                string FileName = string.Format("Language\\{0}.xml", CurrentForm.TypeEx);

                XDocument doc = new XDocument();
                doc.Add(new XElement("FORM", new XAttribute("OLDCaption", CurrentForm.Title), new XAttribute("NEWCaption", CurrentForm.Title)));
                var root = doc.Element("FORM");
                var pb = ADDONBASE._Initializer.SBO_Application.StatusBar.CreateProgressBar("Translation", CurrentForm.Items.Count, true);
                int count = 0;
                Parallel.For(0, CurrentForm.Items.Count, (i) =>
                {
                    // for (int i = 0; i < CurrentForm.Items.Count; i++)
                    // {
                    count++;
                    pb.Value = count;
                    if (CurrentForm.Items.Item(i).Specific is SAPbouiCOM.StaticText)
                    {

                        root.Add(new XElement("ITEM",
                            new XAttribute("UID",
                                CurrentForm.Items.Item(i).UniqueID),
                                new XAttribute("OLDCaption", (CurrentForm.Items.Item(i).Specific as SAPbouiCOM.StaticText).Caption)
                            , new XAttribute("NEWCaption", (CurrentForm.Items.Item(i).Specific as SAPbouiCOM.StaticText).Caption)));

                    }
                    else
                        if (CurrentForm.Items.Item(i).Specific is SAPbouiCOM.Folder)
                    {
                        root.Add(new XElement("ITEM",
                             new XAttribute("UID",
                                 CurrentForm.Items.Item(i).UniqueID),
                                 new XAttribute("OLDCaption", (CurrentForm.Items.Item(i).Specific as SAPbouiCOM.Folder).Caption)
                             , new XAttribute("NEWCaption", (CurrentForm.Items.Item(i).Specific as SAPbouiCOM.Folder).Caption)));

                    }



                    else
                            if (CurrentForm.Items.Item(i).Specific is SAPbouiCOM.Matrix)
                    {
                        var element = new XElement("ITEM",
                            new XAttribute("UID", CurrentForm.Items.Item(i).UniqueID)

                            );
                        var matrix = CurrentForm.Items.Item(i).Specific as SAPbouiCOM.Matrix;
                        for (int j = 0; j < matrix.Columns.Count; j++)
                        {
                            element.Add(new XElement("COLUMN",
                                new XAttribute("UID", matrix.Columns.Item(j).UniqueID.ToString()),
                                new XAttribute("OLDCaption", matrix.Columns.Item(j).TitleObject.Caption)
                            , new XAttribute("NEWCaption", matrix.Columns.Item(j).TitleObject.Caption)

                                ));

                        }
                        root.Add(element);
                    }
                }
                 );
                pb.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pb);
                pb = null;
                GC.Collect();

                doc.Save(FileName);
            });
            tsk.Start();
        }

        public static void SetConditions(this SAPbouiCOM.ChooseFromList cfl, string Alias, SAPbouiCOM.BoConditionOperation Operation, SAPbouiCOM.BoConditionRelationship Relation, string Query)
        {
            var conds = new SAPbouiCOM.Conditions();
            var i = 0;
            var recset = ADDONBASE._Initializer.Company.DoQuery(Query);
            if (recset.RecordCount == 0)
            {
                var cond = conds.Add();
                cond.Alias = Alias;
                cond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                cond.CondVal = "";

            }
            while (!recset.EoF)
            {
                var cond = conds.Add();
                cond.Alias = Alias;
                cond.Operation = Operation;
                cond.CondVal = recset.Fields.Item(Alias).Value.ToString().Trim();
                if (i < recset.RecordCount - 1)
                    cond.Relationship = Relation;
                recset.MoveNext();
                i++;
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
            cfl.SetConditions(conds);
        }
        public static void SetConditions(this SAPbouiCOM.ChooseFromList cfl, string Alias, string Value, SAPbouiCOM.BoConditionOperation Operation)
        {
            SAPbouiCOM.Conditions conds = new SAPbouiCOM.Conditions();
            var cond = conds.Add();
            cond.Alias = Alias;
            cond.Operation = Operation;
            cond.CondVal = Value;
            cfl.SetConditions(conds);
        }
        public static SAPbouiCOM.ChooseFromList GetCFL(this SAPbouiCOM.ChooseFromListEvent cflEventArgs)
        {
            var CurrentForm = _Initializer.SBO_Application.Forms.Item(cflEventArgs.FormUID);
            var cflItem = CurrentForm.ChooseFromLists.Item(cflEventArgs.ChooseFromListUID);
            return cflItem;
        }
        public static XElement CloneElement(this XElement element)
        {
            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n =>
                {
                    XElement e = n as XElement;
                    if (e != null)
                        return CloneElement(e);
                    return n;
                }
                )
            );
        }
        public static void Clear(this SAPbouiCOM.ComboBox thiscmb)
        {
            while (thiscmb.ValidValues.Count > 0)
            {
                thiscmb.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

        }
        //GetIndexAtColum Clear
        public static void Clear(this SAPbouiCOM.ValidValues thiscmb)
        {
            while (thiscmb.Count > 0)
            {
                thiscmb.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

        }
        public static SAPbobsCOM.Recordset DoQuery(this SAPbobsCOM.Company comp, string str, params string[] values)
        {
            var recset = comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            if (values.Count() > 0)
            {
                recset.DoQuery(string.Format(str, values));
                try { File.WriteAllText("DoQuery.sql", string.Format(str, values)); }
                catch { }

            }
            else
            {
                recset.DoQuery(str);
                try { File.WriteAllText("DoQuery.sql", str); }
                catch { }

            }

            return recset;
        }
        internal static bool HasForm(this SAPbouiCOM.Forms Forms, string FormUID)
        {
            var hasvalue = false;
            for (int i = 0; i < Forms.Count; i++)
            {
                if (Forms.Item(i).UniqueID == FormUID)
                    hasvalue = true;
            }
            return hasvalue;
        }
        public static SAPbouiCOM.Item getItem(this SAPbouiCOM.ItemEvent pval)
        {
            return _Initializer.SBO_Application.Forms.Item(pval.FormUID).Items.Item(pval.ItemUID);
        }
        public static SAPbouiCOM.Form getForm(this SAPbouiCOM.ItemEvent pval)
        {
            return _Initializer.SBO_Application.Forms.Item(pval.FormUID);

        }
        public static dynamic getRelevantCellSpecific(this SAPbouiCOM.ItemEvent pval, string ColumnName)
        {
            return (_Initializer.SBO_Application.Forms.Item(pval.FormUID).Items.Item(pval.ItemUID).Specific as SAPbouiCOM.Matrix).GetCellSpecific(ColumnName, pval.Row);
        }
        public static string getUDOID(this SAPbouiCOM.ItemEvent pval)
        {
            string ret = "";
            try
            {
                if (pval.FormTypeEx.Contains('_'))
                {
                    var arr = pval.FormTypeEx.Split('_');
                    ret = arr[arr.Count() - 1];
                }
                else
                    ret = pval.FormTypeEx;
            }
            catch
            {

            }
            return ret;
        }
        public static void ExtractSaveResource(this Assembly a, String filename, String location)
        {
            Stream resFilestream = a.GetManifestResourceStream(filename);
            if (resFilestream != null)
            {
                BinaryReader br = new BinaryReader(resFilestream);
                FileStream fs = new FileStream(location, FileMode.Create); // say 
                BinaryWriter bw = new BinaryWriter(fs);
                byte[] ba = new byte[resFilestream.Length];
                resFilestream.Read(ba, 0, ba.Length);
                bw.Write(ba);
                br.Close();
                bw.Close();
                resFilestream.Close();
            }
            // this.Close(); 
        }
        public static bool Exists(this SAPbouiCOM.DataTable a, string ColumnName, String Value, String Operation = "==")
        {
            bool exists = false;
            var length = a.Rows.Count;
            for (int i = 0; i < length; i++)
            {
                switch (Operation)
                {
                    case "==":
                        {

                            if (a.GetValue(ColumnName, i).ToString() == Value)
                            {
                                exists = true;

                            }
                        }
                        break;
                    case "!=":
                        {

                            if (a.GetValue(ColumnName, i).ToString() != Value)
                            {
                                exists = true;
                            }
                        }
                        break;
                }
            }
            return exists;
        }
        public static bool Exists(this SAPbouiCOM.DBDataSource a, string ColumnName, String Value, String Operation = "==")
        {
            bool exists = false;
            var length = a.Size;
            for (int i = 0; i < length; i++)
            {
                switch (Operation)
                {
                    case "==":
                        {

                            if (a.GetValue(ColumnName, i).ToString().Trim() == Value)
                            {
                                exists = true;
                            }
                        }
                        break;
                    case "!=":
                        {

                            if (a.GetValue(ColumnName, i).ToString().Trim() != Value)
                            {
                                exists = true;
                            }
                        }
                        break;
                }
            }
            return exists;
        }

        public static void ClearAt(this SAPbouiCOM.DataTable a, string ColumnName, String Value, String Operation = "==")
        {
            var length = a.Rows.Count;
            for (int i = 0; i < length; i++)
            {
                switch (Operation)
                {
                    case "==":
                        {

                            if (a.GetValue(ColumnName, i).ToString() == Value)
                            {
                                a.Rows.Remove(i);
                                length = a.Rows.Count;
                                i = i - 1;
                            }
                        }
                        break;
                    case "!=":
                        {

                            if (a.GetValue(ColumnName, i).ToString() != Value)
                            {
                                a.Rows.Remove(i);
                                length = a.Rows.Count;
                                i = i - 1;
                            }
                        }
                        break;
                }
            }
        }
        public static void ClearAt(this SAPbouiCOM.DBDataSource a, string ColumnName, String Value, String Operation = "==")
        {
            var length = a.Size;
            for (int i = 0; i < length; i++)
            {
                switch (Operation)
                {
                    case "==":
                        {

                            if (a.GetValue(ColumnName, i).ToString().Trim() == Value)
                            {
                                a.RemoveRecord(i);
                                length = a.Size;
                                i = i - 1;
                            }
                        }
                        break;
                    case "!=":
                        {

                            if (a.GetValue(ColumnName, i).ToString().Trim() != Value)
                            {
                                a.RemoveRecord(i);
                                length = a.Size;
                                i = i - 1;
                            }
                        }
                        break;
                }
            }
        }
        public static int GetLastIndex(this SAPbouiCOM.DBDataSource a)
        {
            var ind = 0;
            ind = a.Size;
            if (ind > 0) ind = ind - 1;
            return ind;
        }
        public static System.Data.DataTable getDataTable(this SAPbouiCOM.DataTable a)
        {
            System.Data.DataTable tb = new System.Data.DataTable();
            for (int i = 0; i < a.Columns.Count; i++)
            {
                var colm = a.Columns.Item(i);
                tb.Columns.Add(colm.Name.ToString());

            }
            for (int i = 0; i < a.Rows.Count; i++)
            {
                List<object> objs = new List<object>();
                for (int j = 0; j < a.Columns.Count; j++)
                {
                    objs.Add(a.GetValue(j, i));
                }
                tb.Rows.Add(objs.ToArray());
            }
            return tb;
        }
        public static string getDateString(this SAPbouiCOM.UserDataSource ud)
        {
            var recset = _Initializer.Company.DoQuery("Select \"DateFormat\",\"DateSep\" from OADM");
            var dateFormat = Convert.ToString((object)recset.Fields.Item("DateFormat").Value);

            var dateSep = Convert.ToString((object)recset.Fields.Item("DateSep").Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
            if (dateFormat == "0") dateFormat = "DD/MM/YY";
            else if (dateFormat == "1") dateFormat = "DD/MM/YYYY";
            else if (dateFormat == "2") dateFormat = "MM/DD/YY";
            else if (dateFormat == "3") dateFormat = "MM/DD/YYYY";
            else if (dateFormat == "4") dateFormat = "YYYY/MM/DD";
            else if (dateFormat == "5") dateFormat = "DD/Month/YYYY";
            else if (dateFormat == "6") dateFormat = "YY/MM/DD";
            dateFormat = dateFormat.ToLower();
            dateFormat = dateFormat.Replace("mm", "MM");
            //           dateFormat = dateFormat.Replace("/", dateSep);
            var valu = ud.Value.Replace(dateSep, "/");

            return DateTime.ParseExact(valu, dateFormat, null).ToString("yyyyMMdd");
        }
        public static void setDataTable(this SAPbouiCOM.DataTable a, SAPbobsCOM.Recordset tb)
        {
            a.Rows.Clear();
            a.Rows.Add(tb.RecordCount);
            int r = 0;
            while (!tb.EoF)
            {

                for (int c = 0; c < tb.Fields.Count; c++)
                {
                    var columnName = tb.Fields.Item(c).Name;

                    var ccount = a.Columns.Count;

                    for (int ct = 0; ct < ccount; ct++)
                    {
                        var ctxName = a.Columns.Item(ct).Name;
                        if (ctxName.ToLower().Trim() == columnName.ToLower().Trim())
                        {

                            try
                            {
                                a.SetValue(ctxName, r, tb.Fields.Item(columnName).Value);
                            }
                            catch (Exception ex) { ex.PrintString(); }
                            break;
                        }

                    }

                }
                r++;
                tb.MoveNext();
            }
        }

        public static void setDataTable(this SAPbouiCOM.DataTable a, System.Data.DataTable tb)
        {
            a.Rows.Clear();
            a.Rows.Add(tb.Rows.Count);

            for (int r = 0; r < tb.Rows.Count; r++)
            {
                for (int c = 0; c < tb.Columns.Count; c++)
                {
                    var columnName = tb.Columns[c].ColumnName;

                    var ccount = a.Columns.Count;

                    for (int ct = 0; ct < ccount; ct++)
                    {
                        if (a.Columns.Item(ct).Name.Equals(columnName))
                        {

                            try
                            {
                                var k = tb.Rows[r][columnName];
                                if (columnName == "Code")
                                    a.SetValue(columnName, r, Convert.ToString(tb.Rows[r][columnName]));
                                else a.SetValue(columnName, r, tb.Rows[r][columnName]);
                            }
                            catch (Exception ex) { ex.PrintString(); }
                            break;
                        }

                    }

                }
            }
        }

        public static System.Data.DataTable getDataTable(this SAPbouiCOM.DBDataSource a)
        {
            System.Data.DataTable tb = new System.Data.DataTable();
            for (int i = 0; i < a.Fields.Count; i++)
            {
                var colm = a.Fields.Item(i);
                tb.Columns.Add(colm.Name.ToString());

            }
            for (int i = 0; i < a.Size; i++)
            {
                List<object> objs = new List<object>();
                for (int j = 0; j < a.Fields.Count; j++)
                {
                    objs.Add(a.GetValue(j, i));
                }
                tb.Rows.Add(objs.ToArray());
            }
            return tb;
        }
        public static int GetSelectedRow(this SAPbouiCOM.Matrix Matrix)
        {
            int i = 1;
            bool found = false;
            for (i = 1; i <= Matrix.RowCount; i++)
            {
                if (Matrix.IsRowSelected(i))
                {
                    found = true;
                    break;
                }
            }
            if (found)
                return i - 1;
            else return -1;
        }
        public static int getColIDatName(this SAPbouiCOM.Matrix Matrix, string ColUID)
        {
            int i = 0;
            for (i = 0; i < Matrix.Columns.Count; i++)
            {
                if (Matrix.Columns.Item(i).UniqueID == ColUID)
                    break;
            }
            return i;
        }
        public static double getStock(this SAPbobsCOM.Company Company, string itemcode, string whscode, int SerialNumber = -1, string BatchNumber = "")
        {
            var recordset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;

            recordset.DoQuery(string.Format("exec TC_Extensions_GetWHS @itemcode='{0}',@whsCode='{1}', @serialNumber = {2}, @batchnumber ='{3}'", itemcode, whscode, SerialNumber, BatchNumber));
            var value = recordset.Fields.Item(0).Value;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset); GC.Collect();
            return Convert.ToDouble(value);
        }

        public static bool HasColumnName(this SAPbouiCOM.DataColumns Columns, string Name)
        {
            bool tr = false;
            for (int i = 0; i < Columns.Count; i++)
            {
                if (Columns.Item(i).Name == Name)
                {
                    tr = true;
                    break;
                }
            }
            return tr;
        }
        public static object PrintString(this object o)
        {
            Console.WriteLine(o.ToString());
            return o;
        }
        public static int GetLineIdAtValue(this SAPbouiCOM.DBDataSource o, String ColumnName, String Value)
        {
            var ret = -1;
            var length = o.Size;
            for (int i = 0; i < length; i++)
            {
                if (o.GetValue(ColumnName, i).Trim().Equals(Value))
                {
                    ret = i;
                    break;
                }
            }
            return ret;

        }
        public static string GetValueAtLineId(this SAPbouiCOM.DBDataSource o, String Name, String lineid)
        {
            var ret = "";
            var length = o.Size;
            for (int i = 0; i < length; i++)
            {
                if (o.GetValue("LineId", i).Trim().Equals(lineid))
                {
                    ret = o.GetValue(Name, i).Trim();
                    break;
                }
            }
            return ret;
        }
        public static double ToDouble(this string o)
        {
            double value = 0;
            try
            {
                value = Convert.ToDouble(o);
            }
            catch (Exception ex) { }
            return value;
        }
        public delegate bool ActionReturn<in T>(T obj);
        /// <summary>
        /// convert string to double 
        /// </summary>
        /// <param name="o"></param>
        /// <param name="def">Default Value if string is not convertable to double</param>
        /// <returns></returns>
        public static double ToDouble(this string o, double def)
        {
            double value = def;
            try
            {
                value = Convert.ToDouble(o);
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            return value;
        }
        public static int ToInt(this string o)
        {
            int value = 0;
            try
            {
                value = Convert.ToInt32(o);
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            return value;
        }
        /// <summary>
        /// convert string to double 
        /// </summary>
        /// <param name="o"></param>
        /// <param name="def">Default Value if string is not convertable to double</param>
        /// <returns></returns>
        public static int ToInt(this string o, int def)
        {
            int value = def;
            try
            {
                value = Convert.ToInt32(o);
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            return value;
        }
        public static double GETSUM(this SAPbouiCOM.DBDataSource o, object ColumnName)
        {
            var length = o.Size;
            var sum = 0.0;
            for (int i = 0; i < length; i++)
            {
                sum += o.GetValue(ColumnName, i).Trim().ToDouble();
            }
            return sum;
        }
        public static void Foreach(this SAPbouiCOM.DBDataSource o, ActionReturn<int> a)
        {
            var length = o.Size;
            var b = true;
            for (int i = 0; i < length; i++)
            {
                b = a.Invoke(i);
                if (!b) break;
            }
        }
        public static void FillByRecordset(this SAPbouiCOM.ComboBox combo, SAPbobsCOM.Recordset recset)
        {
            while (combo.ValidValues.Count > 0)
                combo.ValidValues.Remove(combo.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
            var code = string.Empty;
            var name = string.Empty;
            while (!recset.EoF)
            {
                code = recset.Fields.Item("code").Value.ToString();
                name = recset.Fields.Item("name").Value.ToString();
                try
                {
                    combo.ValidValues.Add(code, name);
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                recset.MoveNext();
            }

        }
        public static void FillDataTable(this SAPbobsCOM.Recordset recset, System.Data.DataTable table)
        {
            table.Rows.Clear();
            #region Create Columns
            for (int i = 0; i < recset.Fields.Count; i++)
            {
                Type type = null;
                switch (recset.Fields.Item(i).Type)
                {
                    case SAPbobsCOM.BoFieldTypes.db_Alpha:
                        type = typeof(string);
                        break;
                    case SAPbobsCOM.BoFieldTypes.db_Date:
                        type = typeof(DateTime);
                        break;
                    case SAPbobsCOM.BoFieldTypes.db_Float:
                        type = typeof(double);
                        break;
                    case SAPbobsCOM.BoFieldTypes.db_Memo:

                        type = typeof(string);
                        break;
                    case SAPbobsCOM.BoFieldTypes.db_Numeric:
                        type = typeof(float);
                        break;
                    default:
                        break;
                }
                table.Columns.Add(recset.Fields.Item(i).Name, type);
            }

            #endregion
            recset.MoveFirst();
            var count = 0;
            while (!recset.EoF)
            {

                table.Rows.Add();
                for (int i = 0; i < recset.Fields.Count; i++)
                {
                    table.Rows[count][recset.Fields.Item(i).Name] = recset.Fields.Item(i).Value;
                }
                count++;
                recset.MoveNext();
            }

        }
        public static void FillByQuery(this SAPbouiCOM.ComboBox combo, string Query)
        {
            var recset = _Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(Query);
            recset.MoveFirst();
            combo.FillByRecordset(recset);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
            //while (combo.ValidValues.Count > 0)
            //    combo.ValidValues.Remove(combo.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
            //var code = string.Empty;
            //var name = string.Empty;
            //while (!recset.EoF)
            //{
            //    code = recset.Fields.Item("code").Value.ToString();
            //    name = recset.Fields.Item("name").Value.ToString();
            //    try
            //    {
            //        combo.ValidValues.Add(code, name);
            //    }
            //    catch(Exception ex) { ex.AppendInLogFile(); }
            //    recset.MoveNext();
            //}

        }
        public static void FillByQuery(this SAPbouiCOM.ValidValues combo, string Query)
        {
            combo.Clear();
            var recset = _Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(Query);
            recset.MoveFirst();
            combo.FillByRecordset(recset);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset); GC.Collect();
        }
        public static void FillByQuery(this SAPbouiCOM.ValidValues combo, string Query, params object[] objs)
        {
            // combo.Clear();
            var recset = _Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery(string.Format(Query, objs));
            recset.MoveFirst();
            combo.FillByRecordset(recset);

        }

        public static void FillByRecordset(this SAPbouiCOM.ValidValues combo, SAPbobsCOM.Recordset recset)
        {
            while (combo.Count > 0)
                combo.Remove(combo.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
            var code = string.Empty;
            var name = string.Empty;
            while (!recset.EoF)
            {
                code = recset.Fields.Item("Code").Value.ToString();
                name = recset.Fields.Item("Name").Value.ToString();
                try
                {
                    combo.Add(code, name);
                }
                catch (Exception ex) { ex.AppendInLogFile(); }
                recset.MoveNext();
            }

        }
        public static void Fill(this SAPbouiCOM.DBDataSource DetailDS, SAPbouiCOM.DBDataSource DSSource)
        {
            DetailDS.Clear();
            for (int i = 0; i < DSSource.Size; i++)
            {
                DetailDS.InsertRecord(i);
                for (int colid = 0; colid < DSSource.Fields.Count; colid++)
                {
                    var colname = DSSource.Fields.Item(colid).Name;
                    try
                    {

                        DetailDS.SetValue(colname, i, DSSource.GetValue(colname, i).Trim());
                    }
                    catch { }
                }
            }

        }
        public static void FillByRecordset(this SAPbouiCOM.DBDataSource DetailDS, SAPbobsCOM.Recordset recset)
        {
            int j = 0;
            while (DetailDS.Size != 0)
                DetailDS.RemoveRecord(DetailDS.Size - 1);
            try
            {
                recset.MoveFirst();
                while (!recset.EoF)
                {

                    DetailDS.InsertRecord(j);
                    for (int i = 0; i < recset.Fields.Count; i++)
                    {
                        try
                        {
                            if (recset.Fields.Item(i).Value is DateTime)
                                DetailDS.SetValue(recset.Fields.Item(i).Name, j, ((DateTime)recset.Fields.Item(i).Value).ToString("yyyyMMdd"));
                            else
                                DetailDS.SetValue(recset.Fields.Item(i).Name, j, recset.Fields.Item(i).Value.ToString());
                        }
                        catch (Exception ex)
                        {
                            ex.AppendInLogFile();
                        }
                    }
                    j++;
                    recset.MoveNext();
                }

            }
            catch (Exception ex)
            {
                ex.AppendInLogFile();
            }
        }
        /// <summary>
        /// Get Comma Separated String for given ColumnName from DataSet
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="ColumName"></param>
        /// <returns></returns>

        public static string GET_CSV(this SAPbouiCOM.DBDataSource ds, string ColumName)
        {
            string itemcodes = "";
            ds.Foreach((i) =>
            {
                itemcodes += ",'" + ds.GetValue(ColumName, i).Trim() + "'";
                return true;
            });
            if (itemcodes.Length > 0) itemcodes = itemcodes.Substring(1);
            return itemcodes;
        }
        /// <summary>
        /// recordset's field names should be same as that of DetailDS columns
        /// First Column in recordset is KeyColumn
        /// </summary>
        /// <param name="DetailDS"></param>
        /// <param name="recset"></param>
        public static void UpdateByRecordset(this SAPbouiCOM.DBDataSource DetailDS, SAPbobsCOM.Recordset recset)
        {
            int j = 0;
            recset.MoveFirst();
            var keyColumnName = recset.Fields.Item(0).Name;
            var dictionary = recset.GetItemValueDictionary();

            try
            {
                // recset.MoveFirst();
                DetailDS.Foreach((i) =>
                {
                    var key = DetailDS.GetValue(keyColumnName, i).Trim();
                    if (!string.IsNullOrEmpty(key))
                    {
                        var fields = dictionary[key];
                        for (int k = 1; k < fields.Count; k++)
                        {
                            var name = fields[k].Key;
                            var value = fields[k].Value.ToString().Trim();
                            DetailDS.SetValue(name, i, value);
                        }
                    }
                    return true;

                });


            }
            catch (Exception ex)
            {
                ex.AppendInLogFile();
            }
        }
        public static void printAtStatusBar(this Exception ex)
        {
            _Initializer.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        }
        static Dictionary<string, object> Caller = new Dictionary<string, object>();
        const string FileName = "PERFORMANCE_LOG.txt";
        public static void Tic([CallerMemberName] string caller = null)
        {
            if (System.Configuration.ConfigurationManager.AppSettings["ISLOG"].ToString().ToLower() == "true")
            {
                if (Caller.Keys.Contains(caller))
                    Caller.Remove(caller);
                Stopwatch st = Stopwatch.StartNew();
                Caller.Add(caller, st);
                //     var path = System.Environment.CurrentDirectory + FileName;
                //   System.IO.File.AppendAllLines(path, new string[] { "--" + caller + ":TIC--", DateTime.Now.ToString() });
            }
            // look at caller
        }
        public static void Toc([CallerMemberName] string caller = null)
        {

            if (System.Configuration.ConfigurationManager.AppSettings["ISLOG"].ToString().ToLower() == "true")
            {
                var st = Caller[caller] as Stopwatch;
                st.Stop();

                Caller.Remove(caller);
                var path = System.Environment.CurrentDirectory + "//" + FileName;
                //System.IO.File.AppendAllLines(path, new string[] { "--" + caller + ":TOC--", DateTime.Now.ToString(), string.Format("DURATION: {0}", st.Elapsed.ToString()) });

                System.IO.File.AppendAllLines(path, new string[] { string .Format ("--{0}--" ,caller ),
                    string .Format ("Started At= {0}",(DateTime.Now - st.Elapsed).ToString())
                    , string .Format ("Ended AT={0} -- DURATION: {1}",  DateTime.Now.ToString(), st.Elapsed.ToString())
                    , "---------------------" });

            }

        }
        // look at caller
        public static void AppendInLogFile(this Exception ex)
        {
            var path = System.Environment.CurrentDirectory + "\\Error_LOG.txt";
            System.IO.File.AppendAllLines(path, new string[] { "--Error--", DateTime.Now.ToString(), ex.Message, ex.StackTrace, "*-----*----*----*" });
            if (System.Diagnostics.Debugger.IsAttached) Process.Start(path);
        }
        public static void printAtMessageBox(this Exception ex)
        {
            _Initializer.SBO_Application.MessageBox(ex.Message + ex.StackTrace);
        }
        public static void ApplyCFLNameFill(this SAPbouiCOM.EditText a, String DestinationAlias, Func<SAPbouiCOM.DataTable, string> action)
        {
            a.ChooseFromListAfter += (object sboObject, SAPbouiCOM.SBOItemEventArg pVal) =>
            {
                if (a.Item.UniqueID == pVal.ItemUID)
                    try
                    {
                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(true);
                        SAPbouiCOM.SBOChooseFromListEventArg chooseFromListEvent = ((SAPbouiCOM.SBOChooseFromListEventArg)(pVal));
                        var db = _Initializer.SBO_Application.Forms.Item(pVal.FormUID).DataSources.DBDataSources.Item(a.DataBind.TableName);
                        var value = "";
                        if (chooseFromListEvent.SelectedObjects != null)
                        {

                            value = action.Invoke(chooseFromListEvent.SelectedObjects);
                            db.SetValue(DestinationAlias, 0, value);
                        }
                    }
                    catch (Exception ex)
                    {
                        ex.AppendInLogFile();
                    }
                    finally
                    {

                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(false);
                    }

            };

        }

        public static void ApplyCFLNameFill(this SAPbouiCOM.EditText a, String DestinationAlias, string SourceAlias, string Seperator = " ")
        {
            a.ChooseFromListAfter += (object sboObject, SAPbouiCOM.SBOItemEventArg pVal) =>
            {
                if (a.Item.UniqueID == pVal.ItemUID)
                    try
                    {
                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(true);
                        SAPbouiCOM.SBOChooseFromListEventArg chooseFromListEvent = ((SAPbouiCOM.SBOChooseFromListEventArg)(pVal));
                        var db = _Initializer.SBO_Application.Forms.Item(pVal.FormUID).DataSources.DBDataSources.Item(a.DataBind.TableName);
                        var value = "";
                        if (chooseFromListEvent.SelectedObjects != null)
                        {
                            if (!SourceAlias.Contains(','))
                            {
                                value = chooseFromListEvent.SelectedObjects.GetValue(SourceAlias, 0).ToString().Trim();

                            }
                            else
                            {
                                var values = SourceAlias.Split(',');
                                foreach (var item in values)
                                {
                                    value += Seperator + chooseFromListEvent.SelectedObjects.GetValue(item, 0).ToString().Trim();

                                }
                                if (value.Count() > 0) value = value.Substring(Seperator.Count());


                            }

                            db.SetValue(DestinationAlias, 0, value);
                        }
                    }
                    catch (Exception ex)
                    {
                        ex.AppendInLogFile();
                    }
                    finally
                    {

                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(false);
                    }

            };

        }
        public static void ApplyCFLNameFill(this SAPbouiCOM.Column a, String DestinationAlias, string SourceAlias)
        {

            a.ChooseFromListAfter += (object sboObject, SAPbouiCOM.SBOItemEventArg pVal) =>
            {
                ////y a.Cells.Item(pVal.Row).Click();
                SAPbouiCOM.SBOChooseFromListEventArg chooseFromListEvent = ((SAPbouiCOM.SBOChooseFromListEventArg)(pVal));
                if (a.UniqueID == pVal.ColUID && chooseFromListEvent.SelectedObjects != null)
                    try
                    {
                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(true);
                        //following line throws invoke action
                        var db = _Initializer.SBO_Application.Forms.Item(pVal.FormUID).DataSources.DBDataSources.Item(a.DataBind.TableName);
                        var value = chooseFromListEvent.SelectedObjects.GetValue(SourceAlias, 0);
                        var strValue = value.ToString().Trim();
                        if (value is DateTime) strValue = ((DateTime)value).ToString("yyyyMMdd");

                        db.SetValue(DestinationAlias, pVal.Row - 1, strValue);
                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(false);
                        //(_Initializer.SBO_Application.Forms.Item(pVal.FormUID).Items.Item(pVal.ItemUID).Specific as SAPbouiCOM.Matrix).LoadFromDataSourceEx();
                    }
                    catch (Exception ex)
                    {
                        ex.AppendInLogFile();
                    }
            };

        }
        public static void ApplyCFLNameFill(this SAPbouiCOM.Column a, String DestinationAlias, Func<SAPbouiCOM.SBOItemEventArg, String> GetValueFunction)
        {

            a.ChooseFromListAfter += (object sboObject, SAPbouiCOM.SBOItemEventArg pVal) =>
            {

               // a.Cells.Item(pVal.Row).Click();
               SAPbouiCOM.SBOChooseFromListEventArg chooseFromListEvent = ((SAPbouiCOM.SBOChooseFromListEventArg)(pVal));
                if (a.UniqueID == pVal.ColUID && chooseFromListEvent.SelectedObjects != null)
                    try
                    {
                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(true);
                        var db = _Initializer.SBO_Application.Forms.Item(pVal.FormUID).DataSources.DBDataSources.Item(a.DataBind.TableName);

                        var strValue = GetValueFunction(pVal);

                        db.SetValue(DestinationAlias, pVal.Row - 1, strValue);
                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(false);
                       //(_Initializer.SBO_Application.Forms.Item(pVal.FormUID).Items.Item(pVal.ItemUID).Specific as SAPbouiCOM.Matrix).LoadFromDataSourceEx();
                   }
                    catch (Exception ex)
                    {
                        ex.AppendInLogFile();
                    }
            };

        }


        public static string CreateUserQuery(this SAPbobsCOM.Company Company, string QueryDescription, string Query)
        {
            string ans = "";
            try
            {
                SAPbobsCOM.UserQueries oQuery = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries) as SAPbobsCOM.UserQueries;

                oQuery.Query = Query;

                oQuery.QueryCategory = -1;

                oQuery.QueryDescription = QueryDescription;

                if (oQuery.Add() != 0)
                {
                    ans = Company.GetNewObjectKey();
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oQuery);
            }
            catch (Exception ex) { ex.AppendInLogFile(); }
            return ans;
        }

        public static void ApplyCFLNameFillByUID(this SAPbouiCOM.EditText a, String DestinationID, string SourceAlias)
        {
            a.ChooseFromListAfter += (object sboObject, SAPbouiCOM.SBOItemEventArg pVal) =>
            {
                if (a.Item.UniqueID == pVal.ItemUID)
                    try
                    {
                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(true);
                        SAPbouiCOM.SBOChooseFromListEventArg chooseFromListEvent = ((SAPbouiCOM.SBOChooseFromListEventArg)(pVal));
                        var value = chooseFromListEvent.SelectedObjects.GetValue(SourceAlias, 0).ToString().Trim();
                        (_Initializer.SBO_Application.Forms.Item(pVal.FormUID).Items.Item(DestinationID).Specific as SAPbouiCOM.EditText).Value = value;

                    }
                    catch (Exception ex)
                    {
                        ex.AppendInLogFile();
                    }
                    finally
                    {

                        _Initializer.SBO_Application.Forms.Item(pVal.FormUID).Freeze(false);
                    }

            };

        }
        public static string List2CSV(this List<String> a)
        {
            var val = "";
            foreach (var item in a)
            {
                val += "," + string.Format("'{0}'", item);
            }
            if (!String.IsNullOrEmpty(val)) val = val.Substring(1);
            return val;
        }
        public static Dictionary<string, List<KeyValuePair<string, object>>> GetItemValueDictionary(this SAPbobsCOM.Recordset recset)
        {

            var keycolumnName = recset.Fields.Item(0).Name;
            var itemquantityDictionary = new Dictionary<string, List<KeyValuePair<string, object>>>();
            // var recset = GetRecordSet(query);
            while (!recset.EoF)
            {
                var itemcode = recset.Fields.Item(keycolumnName).Value.ToString();
                var lst = new List<KeyValuePair<string, object>>();
                for (int i = 0; i < recset.Fields.Count; i++)
                {
                    lst.Add(new KeyValuePair<string, object>(recset.Fields.Item(i).Name, recset.Fields.Item(i).Value));
                }

                var quantity = recset.Fields.Item(1).Value.ToString().ToDouble();
                if (itemquantityDictionary.Keys.Contains(itemcode))
                {
                    itemquantityDictionary[itemcode] = lst;
                }
                else
                {
                    itemquantityDictionary.Add(itemcode, lst);
                }

                recset.MoveNext();
            }
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
            return itemquantityDictionary;
        }

    }
}