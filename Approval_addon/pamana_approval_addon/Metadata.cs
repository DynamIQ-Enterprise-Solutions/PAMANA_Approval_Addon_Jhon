using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace pamana_approval_addon
{
    class Metadata : B1Provider
    {
        public static void CreateDefaultObjects()
        {
            CreateTables("DYN_EMAIL_CONFIG", "Email Config", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            InsertDefVal();
        }

        private static void InsertDefVal()
        {
            SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            SAPbobsCOM.Recordset oCheckRows = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);


            if (TableExist("DYN_EMAIL_CONFIG"))
            {
                oCheckRows.DoQuery("SELECT * FROM [@DYN_EMAIL_CONFIG]");

                if (oCheckRows.RecordCount == 0)
                {
                    try
                    {

                        SAPbobsCOM.Recordset oInsertValues = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                        string query = "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'email_receiver','freelancon90210@gmail.com' " + System.Environment.NewLine;
                        query += "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'email_sender','ragnaangar1000@gmail.com' " + System.Environment.NewLine;
                        query += "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'email_sender_password','cssjoprrvtftddja' " + System.Environment.NewLine;
                        query += "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'port','587' " + System.Environment.NewLine;
                        query += "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'smtp','smtp.gmail.com' " + System.Environment.NewLine;
                        query += "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'report_name','PR' " + System.Environment.NewLine;
                        query += "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'crystal_user','sa' " + System.Environment.NewLine;
                        query += "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'crystal_pass','1234' " + System.Environment.NewLine;
                        query += "INSERT INTO \"@DYN_EMAIL_CONFIG\" SELECT 'attachment_path','C:\\RPT'";

                        oInsertValues.DoQuery(query);
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.StatusBar.SetText("DEFAULT VALUES CREATED SUCCESFULY", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }


            }
        }
    }
}
