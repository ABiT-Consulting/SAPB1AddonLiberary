using ADDONBASE.Extensions;
using SAPbobsCOM;
using System;
namespace ADDONBASE
{

    public class UserDefinedObject : IDisposable
    {
        private string _UDO;
        SAPbobsCOM.Company oCompany
        {
            get
            {
                return _Initializer.Company;
            }
        }
        SAPbobsCOM.CompanyService oCompService;
        private SAPbobsCOM.BoUDOObjType UDOType;

        SAPbobsCOM.GeneralService oClassSubjectsGeneralService;
        SAPbobsCOM.GeneralData oClassSubjectsHeaderGeneralData;

        SAPbobsCOM.GeneralDataParams oGenralParameter;

        public UserDefinedObject(string _UDO, SAPbobsCOM.BoUDOObjType UDOType)
        {
            this._UDO = _UDO;
            this.UDOType = UDOType;

            oCompService = oCompany.GetCompanyService();
            //  if (!oCompany.InTransaction)// oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            // oCompany.StartTransaction();

            oClassSubjectsGeneralService = (SAPbobsCOM.GeneralService)oCompService.GetGeneralService(_UDO);
        }
        public void Commit()
        {
            if (oCompany.InTransaction)
            {
                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
        }
        string ObjectKey;
        public void GetByKey(String ObjectKey)
        {
            this.ObjectKey = ObjectKey;
            switch (UDOType)
            {
                case SAPbobsCOM.BoUDOObjType.boud_Document:
                    oGenralParameter = (SAPbobsCOM.GeneralDataParams)oClassSubjectsGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGenralParameter.SetProperty("DocEntry", ObjectKey.Trim());
                    try
                    {
                        oClassSubjectsHeaderGeneralData = oClassSubjectsGeneralService.GetByParams(oGenralParameter);

                    }
                    catch (Exception ex) { ex.AppendInLogFile(); }
                    break;
                case SAPbobsCOM.BoUDOObjType.boud_MasterData:
                    oGenralParameter = (SAPbobsCOM.GeneralDataParams)oClassSubjectsGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGenralParameter.SetProperty("Code", ObjectKey.Trim());

                    oClassSubjectsHeaderGeneralData = oClassSubjectsGeneralService.GetByParams(oGenralParameter);
                    break;
            }
        }
        public dynamic GetProperty(String Name)
        {
            if (oClassSubjectsHeaderGeneralData == null)
                oClassSubjectsHeaderGeneralData = (SAPbobsCOM.GeneralData)oClassSubjectsGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
            return oClassSubjectsHeaderGeneralData.GetProperty(Name);

        }
        public dynamic GetProperty(String Name, int rowid, string datatable)
        {
            object o = null;
            var child = oClassSubjectsHeaderGeneralData.Child(datatable);
            if (child.Count > rowid)
            {
                var DataRow = child.Item(rowid);
                o = DataRow.GetProperty(Name);
            }

            return o;
        }
        public void SetProperty(String Name, Object Value)
        {
            if (oClassSubjectsHeaderGeneralData == null)
                oClassSubjectsHeaderGeneralData = (SAPbobsCOM.GeneralData)oClassSubjectsGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);




            oClassSubjectsHeaderGeneralData.SetProperty(Name, Value);
        }
        public void RemoveRow(int rowid, string datatable)
        {
            var child = oClassSubjectsHeaderGeneralData.Child(datatable);
            if (child.Count > rowid)
            {
                child.Remove(rowid);
            }
        }
        public void SetProperty(String Name, Object Value, int rowid, string datatable)
        {
            var child = oClassSubjectsHeaderGeneralData.Child(datatable);
            if (child.Count > rowid)
            {
                var DataRow = child.Item(rowid);
                DataRow.SetProperty(Name, Value);
            }
            else
            {
                var DataRow = child.Add();
                DataRow.SetProperty(Name, Value);
            }
        }
        public void Update()
        {
            oClassSubjectsGeneralService.Update(oClassSubjectsHeaderGeneralData);
        }
        public void Add()
        {
            var param = oClassSubjectsGeneralService.Add(oClassSubjectsHeaderGeneralData);

        }
        public void Close()
        {

            oClassSubjectsGeneralService.Close(oGenralParameter);

        }

        public int GetSize(string datatable)
        {
            var child = oClassSubjectsHeaderGeneralData.Child(datatable);
            return child.Count;
        }

        public void Dispose()
        {

        }
    }
}
