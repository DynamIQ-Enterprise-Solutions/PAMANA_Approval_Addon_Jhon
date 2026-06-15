using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System.IO;

namespace pamana_approval_addon
{
    class Attachmentss
    {
        private static SAPbobsCOM.Company oCompany { get; set; }
        public Attachmentss()
        {
        }

        public static void GeneratePDF(string reportPath, string pdfOutputPath, string docEntry, string report_name)
        {
            // Create an instance of ReportDocument
            ReportDocument reportDocument = new ReportDocument();
            string output_file = string.Empty;
            try
            {              
                oCompany = new SAPbobsCOM.Company();
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                SAPbobsCOM.Recordset oGetReportCredentials = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                
                string temp_path = Path.GetDirectoryName(reportPath);
                string temp_folder_name = Path.GetFileName(reportPath.TrimEnd(Path.DirectorySeparatorChar));
                string cr_report_cred_user = "";
                string cr_report_cred_pass = "";

                oGetReportCredentials.DoQuery("select * from [@DYN_EMAIL_CONFIG] where code = 'crystal_user' ");

                if (oGetReportCredentials.RecordCount > 0)
                {
                    cr_report_cred_user = oGetReportCredentials.Fields.Item("Name").Value.ToString();
                }

                oGetReportCredentials.DoQuery("select * from [@DYN_EMAIL_CONFIG] where code = 'crystal_pass' ");

                if (oGetReportCredentials.RecordCount > 0)
                {
                    cr_report_cred_pass = oGetReportCredentials.Fields.Item("Name").Value.ToString();
                }

                if (!Directory.Exists(reportPath))
                {
                    Directory.CreateDirectory(temp_path + "\\" + temp_folder_name);
                    File.Copy(Directory.GetCurrentDirectory() + "\\REPORT" + $"\\{report_name}.rpt", temp_path + "\\" + temp_folder_name + $"\\{report_name}.rpt");
                }

                // Load the Crystal Report
                string asdad = reportPath + $"\\{report_name}" + ".rpt";
                reportDocument.Load(reportPath + $"\\{report_name}" + ".rpt");

                // If your report connects to a database, set up the connection information.
                // Adjust these values to match your SAP Business One database settings.
                ConnectionInfo connectionInfo = new ConnectionInfo
                {
                    ServerName = oCompany.Server,
                    DatabaseName = oCompany.CompanyDB,
                    UserID = cr_report_cred_user,//EncryptDecrypt.Decrypt(cr_report_cred_user, "pamanapraddon"),
                    Password = cr_report_cred_pass
                };

                // Apply the connection info to each table used in the report.
                foreach (Table table in reportDocument.Database.Tables)
                {
                    TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                    tableLogOnInfo.ConnectionInfo = connectionInfo;
                    table.ApplyLogOnInfo(tableLogOnInfo);
                }

                // Set any parameters required by the report.
                // For example, if your report expects a parameter named "DocEntry":
                reportDocument.SetParameterValue("DocKey@", docEntry);

                // Set up the export options for PDF format.
                DiskFileDestinationOptions diskOptions = new DiskFileDestinationOptions
                {
                    DiskFileName = pdfOutputPath + "\\PR" + docEntry.ToString() + ".pdf"

                };

                output_file = "\\PR" + docEntry.ToString() + ".pdf";
                ExportOptions exportOptions = reportDocument.ExportOptions;
                exportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                exportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                exportOptions.DestinationOptions = diskOptions;

                // Export the report to PDF.
                reportDocument.Export();

                if (!Directory.Exists(pdfOutputPath + "\\SUCCESS"))
                {
                    Directory.CreateDirectory(pdfOutputPath + "\\SUCCESS");
                }

                //Console.WriteLine("PDF generated successfully at: " + pdfOutputPath);
                Application.SBO_Application.StatusBar.SetText("PDF generated successfully at: " + pdfOutputPath + "\\SUCCESS" + output_file, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                // Handle any errors that might occur during report load or export.
                Logger.WriteToFile("ERROR", "Attachment.GeneratePDF", ex.Message);
                Logger.WriteToFile("ERROR", "Attachment.GeneratePDF", ex.Message);
                Application.SBO_Application.StatusBar.SetText("Error generating PDF: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                // Clean up resources.
                reportDocument.Close();
                reportDocument.Dispose();

                System.IO.File.Copy(pdfOutputPath + output_file, pdfOutputPath + "\\SUCCESS" + output_file, true);
                System.IO.File.Delete(pdfOutputPath + output_file);
            }
        }
    }
}
