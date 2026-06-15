using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using NetMail = System.Net.Mail;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace pamana_approval_addon
{
    class Email : Program
    {
        public static void CreateEmail(string message, string attachment_path = "C:\\RPT\\SUCCESS\\", string attachment_name = "")
        {

            //ORIG
            //string smtp = string.Empty;
            //int port = 0;
            //string sender_email = string.Empty;
            //string sender_password = string.Empty;
            //string receiver_email = string.Empty;

            //SAPbobsCOM.Recordset oGetEmailInfo = (SAPbobsCOM.Recordset)oGlobalCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                //[15JUN2026] NEW PPOSITION
                string smtp = string.Empty;
                int port = 0;
                string sender_email = string.Empty;
                string sender_password = string.Empty;
                string receiver_email = string.Empty;

                SAPbobsCOM.Recordset oGetEmailInfo = (SAPbobsCOM.Recordset)oGlobalCompany.GetBusinessObject(BoObjectTypes.BoRecordset);


                oGetEmailInfo.DoQuery("select * from [@DYN_EMAIL_CONFIG]");

                if (oGetEmailInfo.RecordCount > 0)
                {
                    oGetEmailInfo.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'port' ");
                    port = int.Parse(oGetEmailInfo.Fields.Item("Name").Value.ToString());

                    oGetEmailInfo.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'smtp' ");
                    smtp = oGetEmailInfo.Fields.Item("Name").Value.ToString();

                    oGetEmailInfo.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'email_sender' ");
                    sender_email = oGetEmailInfo.Fields.Item("Name").Value.ToString();

                    oGetEmailInfo.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'email_sender_password' ");
                    sender_password = oGetEmailInfo.Fields.Item("Name").Value.ToString();

                    oGetEmailInfo.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'email_receiver' ");
                    receiver_email = oGetEmailInfo.Fields.Item("Name").Value.ToString();
                }
                var fromAddress = new NetMail.MailAddress(sender_email, "SAP Approval Addon Mailer");
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                var smtpClient = new NetMail.SmtpClient
                {
                    Host = smtp,
                    Port = port,
                    EnableSsl = true,
                    DeliveryMethod = NetMail.SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(sender_email, sender_password),
                };

                var mailmessage = new NetMail.MailMessage
                {
                    From = fromAddress,
                    Subject = "SAP Approval Addon - Approved PR's",
                    Body = message,
                };

                foreach (var address in receiver_email.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mailmessage.To.Add(address);
                }

                NetMail.Attachment data = new NetMail.Attachment(attachment_path + "\\SUCCESS" + "\\" + attachment_name);

                mailmessage.Attachments.Add(data);
                smtpClient.Send(mailmessage);

                data.Dispose();

                //[15JUN2026]
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetEmailInfo);
                GC.Collect();


            }
            catch (Exception ex)
            {
                Logger.WriteToFile("ERROR", "CreateEmail", ex.Message);
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            //finally
            //{
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetEmailInfo);
            //    GC.Collect();
            //}

        }
    }
}
