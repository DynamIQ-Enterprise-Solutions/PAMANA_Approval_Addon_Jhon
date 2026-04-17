using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace pamana_approval_addon
{
    class B1Provider
    {
        public static int CreateTables(string tableName, string tableDescription, SAPbobsCOM.BoUTBTableType tableType)
        {
            SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            int lRetCode = -1;

            string sErrMsg;

            //****************************************************************************
            // Add Header and Row tables
            //****************************************************************************
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));

            bool exist = TableExist(tableName);
            if (exist == false)
            {
                oUserTablesMD.TableName = tableName;
                oUserTablesMD.TableDescription = tableDescription;
                oUserTablesMD.TableType = tableType;
                // Add the table
                lRetCode = oUserTablesMD.Add();

                // check for errors in the process
                if (lRetCode != 0)
                {
                    //LogProvider.LogTableCreation("Add", tableName, false, oGlobalCompany.GetLastErrorDescription());
                    //Logger.WriteToFile("ERROR", "B1Provider.CreateTables", tableName + " - " + oGlobalCompany.GetLastErrorDescription());
                }
                else
                {
                    //LogProvider.LogTableCreation("Add", tableName, true, "");
                    //Logger.WriteToFile("OK", "B1Provider.CreateTables", tableName);
                    //sendMessageToStatusBar("OK", $"Table {tableName} was succesfuly created.");
                }
            }
            else
            {
                // oUserTablesMD.TableDescription = tableDescription;
                // Add the table
                lRetCode = 0;
                //LogProvider.LogTableCreation("Update", tableName, true, "");
                //Logger.WriteToFile("UPDATE", "B1Provider.CreateTables", tableName);
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
            GC.Collect(); // Release the handle to the User Fields

            return lRetCode;

        }

        public static bool TableExist(string TableName)
        {
            SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            bool boolIdent = false;
            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
            boolIdent = oUserTablesMD.GetByKey(TableName);


            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
            GC.Collect(); // Release the handle to the User Fields

            return (boolIdent);
        }
    }
}
