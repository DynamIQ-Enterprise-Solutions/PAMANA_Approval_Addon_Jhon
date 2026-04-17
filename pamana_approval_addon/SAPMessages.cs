using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace pamana_approval_addon
{
    class SAPMessages: Program
    {
        public static void sendAlertsToUser(string msg, string sapuser, string draftentry)
        {
            SAPbobsCOM.Messages oMsg = (SAPbobsCOM.Messages)oGlobalCompany.GetBusinessObject(BoObjectTypes.oMessages);

            try
            {
                oMsg.MessageText = msg;
                oMsg.Subject = "Purchase Request error upon Approval.";

                oMsg.Recipients.Add();
                oMsg.Recipients.SetCurrentLine(0);

                oMsg.Recipients.UserCode = sapuser;
                oMsg.Recipients.NameTo = sapuser;
                oMsg.Recipients.SendInternal = BoYesNoEnum.tYES;
                oMsg.AddDataColumn("Check PR Draft", "Check PR Draft", BoObjectTypes.oDrafts, draftentry);
                oMsg.Add();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText("ADDON ERROR.sendAlertsToUser " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
