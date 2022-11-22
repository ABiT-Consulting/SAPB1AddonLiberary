using System;

namespace PLDAdder
{
    internal class PLDAdder
    {

        SAPbouiCOM.Application oApplication
        {
            get
            {
                return ADDONBASE._Initializer.SBO_Application;
            }
        }
        SAPbobsCOM.Company oCompany
        {
            get
            {
                return ADDONBASE._Initializer.Company;
            }
        }
        internal PLDAdder()
        {

        }
        String addReportType(String typeName, String AddonName, String UDOName, String menuId)
        {
            SAPbobsCOM.ReportTypesService rptTypeService;
            SAPbobsCOM.ReportType newType;
            SAPbobsCOM.ReportTypeParams newTypeParam;
            rptTypeService = oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService) as SAPbobsCOM.ReportTypesService;
            newType = rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType) as SAPbobsCOM.ReportType;
            var addReportType = "";

            try
            {

                GC.Collect();
                // newType.TypeCode = typeCode
                newType.TypeName = typeName;
                newType.AddonName = AddonName;
                newType.AddonFormType = UDOName;
                newType.MenuID = menuId;
                newTypeParam = rptTypeService.AddReportType(newType);
                addReportType = oCompany.GetNewObjectKey();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(rptTypeService);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newType);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newTypeParam);
            }
            catch (Exception ex)
            {
                oApplication.SetStatusBarMessage(ex.Message);
            }

            return addReportType;

        }

        internal String getReportTypeCode(String typeName, String udoName, String addonName, String menuId)
        {
            String ReportTypeCode = "";
            try
            {

                SAPbobsCOM.Recordset rsetField = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                var oFlag = "";
                var s = "SELECT top 1 \"CODE\", \"NAME\", \"DEFLT_REP\", \"ADD_NAME\", \"FRM_TYPE\", \"MNU_ID\", \"IS_SYS\", \"DEFLT_SEQ\", \"TYPE\"   FROM  \"RTYP\"  where \"NAME\"  = '{0}' and  \"FRM_TYPE\" ='{1}' and  \"ADD_NAME\" ='{2}' ";
                //"SELECT     top 1 CODE, NAME, DEFLT_REP, ADD_NAME, FRM_TYPE, MNU_ID, IS_SYS, DEFLT_SEQ, TYPE   FROM RTYP where NAME = '{0}' and FRM_TYPE='{1}' and ADD_NAME='{2}'";
                s = string.Format(s, typeName, udoName, addonName);
                //     s = "SELECT     top 1 CODE, NAME, DEFLT_REP, ADD_NAME, FRM_TYPE, MNU_ID, IS_SYS, DEFLT_SEQ, TYPE   FROM RTYP where NAME = '" + typeName + "' and FRM_TYPE='" + udoName + "' and ADD_NAME='" + addonName + "'";
                rsetField.DoQuery(s);
                if (rsetField.EoF)
                    ReportTypeCode = addReportType(typeName, addonName, udoName, menuId);
                else
                    ReportTypeCode = rsetField.Fields.Item("CODE").Value.ToString();



                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsetField);
                rsetField = null;
                GC.Collect();
                return ReportTypeCode;
            }
            catch (Exception ex)
            {
                oApplication.StatusBar.SetText("Failed to Column Exists : " + ex.Message);
            }
            finally
            {
            }
            return ReportTypeCode;


        }
    }
}
