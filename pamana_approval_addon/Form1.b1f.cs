using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using System.IO;


namespace pamana_approval_addon
{
    [FormAttribute("2100100720", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {

            LoadGrid();

        }

        private void LoadGrid()
        {

            //ORIG POSTION
            //SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

            //oUser.GetByKey(Program.oGlobalCompany.UserSignature);

            //this.UIAPIRawForm.Freeze(true);
            
            try
            {
                //NEW POSTION   [08MAY2026]  
                SAPbobsCOM.Users oUser = (SAPbobsCOM.Users)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

                oUser.GetByKey(Program.oGlobalCompany.UserSignature);

                this.UIAPIRawForm.Freeze(true);



                string query = $"select CAST('' as varchar(20)) as 'Decision', ODRF.DocNum 'Doc. Draft No.', OWDD.DraftEntry 'Draft Key', Originator.U_NAME 'Originator', OWDD.Remarks 'Draft Remarks', OWDD.WddCode 'Approval Code' " +
                           $"from OWDD left join WDD1 on OWDD.WddCode = WDD1.WddCode left join OWTM on OWTM.WtmCode = OWDD.WtmCode left join WTM2 on WTM2.WtmCode = OWTM.WtmCode left join WST1 on WST1.WstCode = WTM2.WstCode " +
                           $"left join OUSR Approver on Approver.USERID = WST1.UserID left join ODRF on ODRF.DocEntry = OWDD.DraftEntry left join OUSR Originator on Originator.USERID = OWDD.OwnerID " +
                           $"where WDD1.Status = 'W' and OWDD.ObjType = '1470000113' and Approver.USER_CODE = '{oUser.UserCode}' and OWDD.ProcesStat != 'C' ";
                this.Grid0.DataTable.ExecuteQuery(query);



                SAPbouiCOM.ComboBoxColumn oCmbDecision;
                SAPbouiCOM.EditTextColumn oDraftKey;

                Grid0.Columns.Item("Draft Key").Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                oDraftKey = (SAPbouiCOM.EditTextColumn)Grid0.Columns.Item("Draft Key");
                oDraftKey.LinkedObjectType = "112";

                Grid0.Columns.Item("Decision").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                oCmbDecision = (SAPbouiCOM.ComboBoxColumn)Grid0.Columns.Item("Decision");
                oCmbDecision.ValidValues.Add("", "");
                oCmbDecision.ValidValues.Add("Approve", "Approve");
                oCmbDecision.ValidValues.Add("Reject", "Reject");

                for (int i = 0; i < Grid0.Columns.Count; i++)
                {
                    if (Grid0.Columns.Item(i).TitleObject.Caption != "Decision")
                    {
                        Grid0.Columns.Item(i).Editable = false;
                    }
                }
            }
            catch (Exception ex)
            {
                this.UIAPIRawForm.Freeze(false);

                Logger.WriteToFile("ERROR", "LoadGrid", ex.Message);     //ADD ERROR LOG [17APR2026]
                Application.SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); //ADD ERROR LOG [17APR2026]
                
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
            }

        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            try
            {
                this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));       //CAUSE ERROR   [08MAY2026] Seen By ARLYN
                this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
                this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
                this.Button0.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Button0_PressedBefore);
                this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
                this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
                this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
                this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
                this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
                this.OnCustomInitialize();
            }
            catch (Exception ex)
            {
                this.UIAPIRawForm.Freeze(false);

                Logger.WriteToFile("ERROR", "OnInitializeComponent", ex.Message);     //ADD ERROR LOG [08MAY2026]
                Application.SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); //ADD ERROR LOG [17APR2026]               
            }
           
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Button Button0;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                MessageForm.displayStatus("Processing decisions... please wait... ", Application.SBO_Application.Forms.ActiveForm);

                for (int i = 0; i < Grid0.Rows.Count; i++)
                {

                    if (Grid0.DataTable.GetValue("Decision", Grid0.GetDataTableRowIndex(i)).ToString() == "Approve")
                    {
                        ApproveDocument(int.Parse(Grid0.DataTable.GetValue("Approval Code", Grid0.GetDataTableRowIndex(i)).ToString()), int.Parse(Grid0.DataTable.GetValue("Draft Key", Grid0.GetDataTableRowIndex(i)).ToString()));
                    }
                    else if (Grid0.DataTable.GetValue("Decision", Grid0.GetDataTableRowIndex(i)).ToString() == "Reject")
                    {
                        RejectDocument(int.Parse(Grid0.DataTable.GetValue("Approval Code", Grid0.GetDataTableRowIndex(i)).ToString()));
                    }
                }

                MessageForm.hideStatus();
            }
            catch (Exception ex)
            {
                 Logger.WriteToFile("ERROR", "Button0_PressedAfter", ex.Message);   //ADD ERROR LOG [29MAR2026]
                Application.SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); //ADD ERROR LOG [29MAR2026]
            }
            finally
            {
                LoadGrid();
            }

        }

        private void ApproveDocument(int WddCode, int draftKey)
        {

            //---ORIG POSITION
            //SAPbobsCOM.ApprovalRequestsService oApprovalRequestsService = (SAPbobsCOM.ApprovalRequestsService)Program.oGlobalCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService);
            //SAPbobsCOM.ApprovalRequest oApprovalRequest = (SAPbobsCOM.ApprovalRequest)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequest);
            //SAPbobsCOM.ApprovalRequestParams oApprovalRequestParams = (SAPbobsCOM.ApprovalRequestParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);
            //SAPbobsCOM.ApprovalRequestsParams oApprovalRequestsParams = (SAPbobsCOM.ApprovalRequestsParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestsParams);
            //SAPbobsCOM.ApprovalRequestDecision oApprovalRequestDecision;

            try
            {
                //[07MAY2026]
                SAPbobsCOM.ApprovalRequestsService oApprovalRequestsService = (SAPbobsCOM.ApprovalRequestsService)Program.oGlobalCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService);
                SAPbobsCOM.ApprovalRequest oApprovalRequest = (SAPbobsCOM.ApprovalRequest)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequest);
                SAPbobsCOM.ApprovalRequestParams oApprovalRequestParams = (SAPbobsCOM.ApprovalRequestParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);
                SAPbobsCOM.ApprovalRequestsParams oApprovalRequestsParams = (SAPbobsCOM.ApprovalRequestsParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestsParams);
                SAPbobsCOM.ApprovalRequestDecision oApprovalRequestDecision;

                
                oApprovalRequestsParams = oApprovalRequestsService.GetAllApprovalRequestsList();
                oApprovalRequestParams.Code = WddCode;
                oApprovalRequest = oApprovalRequestsService.GetApprovalRequest(oApprovalRequestParams);

                if (oApprovalRequest.ObjectType == "1470000113")
                {
                    oApprovalRequestDecision = oApprovalRequest.ApprovalRequestDecisions.Add();
                    oApprovalRequestDecision.Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardApproved;
                    oApprovalRequestDecision.ApproverUserName = Program.oGlobalCompany.UserName;
                    oApprovalRequestDecision.ApproverPassword = this.EditText0.Value;
                    oApprovalRequestDecision.Remarks = "Approved by: PR Approver Addon";
                    oApprovalRequestsService.UpdateRequest(oApprovalRequest);
                }

                //SaveDraftToDoc(draftKey);  //ORIG

                SaveDraftToDoc_2(draftKey);  //[20MAY2026]



                //[07MAY2026]
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestsService);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequest);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestParams);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestsParams);

                GC.Collect();

            }
            catch (Exception ex)
            {
                Logger.WriteToFile("ERROR", "ApproveDocument", ex.Message);
                Application.SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            //finally
            //{
                //---ORIG POSITION
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestsService);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequest);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestParams);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestsParams);
                //GC.Collect();
            //}
        }

        private void RejectDocument(int WddCode)
        {

            //--ORIG POSITION
            //SAPbobsCOM.ApprovalRequestsService oApprovalRequestsService = (SAPbobsCOM.ApprovalRequestsService)Program.oGlobalCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService);
            //SAPbobsCOM.ApprovalRequest oApprovalRequest = (SAPbobsCOM.ApprovalRequest)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequest);
            //SAPbobsCOM.ApprovalRequestParams oApprovalRequestParams = (SAPbobsCOM.ApprovalRequestParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);
            //SAPbobsCOM.ApprovalRequestsParams oApprovalRequestsParams = (SAPbobsCOM.ApprovalRequestsParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestsParams);
            //SAPbobsCOM.ApprovalRequestDecision oApprovalRequestDecision;

            try
            {   

                //[07MAY2026] ------
                SAPbobsCOM.ApprovalRequestsService oApprovalRequestsService = (SAPbobsCOM.ApprovalRequestsService)Program.oGlobalCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService);
                SAPbobsCOM.ApprovalRequest oApprovalRequest = (SAPbobsCOM.ApprovalRequest)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequest);
                SAPbobsCOM.ApprovalRequestParams oApprovalRequestParams = (SAPbobsCOM.ApprovalRequestParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);
                SAPbobsCOM.ApprovalRequestsParams oApprovalRequestsParams = (SAPbobsCOM.ApprovalRequestsParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestsParams);
                SAPbobsCOM.ApprovalRequestDecision oApprovalRequestDecision;


                oApprovalRequestsParams = oApprovalRequestsService.GetAllApprovalRequestsList();
                oApprovalRequestParams.Code = WddCode;

                oApprovalRequest = oApprovalRequestsService.GetApprovalRequest(oApprovalRequestParams);

                if (oApprovalRequest.ObjectType == "1470000113")
                {
                    oApprovalRequestDecision = oApprovalRequest.ApprovalRequestDecisions.Add();
                    oApprovalRequestDecision.Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardNotApproved;
                    oApprovalRequestDecision.ApproverUserName = Program.oGlobalCompany.UserName;
                    oApprovalRequestDecision.ApproverPassword = this.EditText0.Value;
                    oApprovalRequestDecision.Remarks = "Rejected by PR Approver Addon";
                    oApprovalRequestsService.UpdateRequest(oApprovalRequest);
                }

                //[07MAY2026] ------
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestsService);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequest);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestParams);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestsParams);
                GC.Collect();
            }
            catch (Exception ex)
            {
                Logger.WriteToFile("ERROR", "RejectDocument", ex.Message);
                Application.SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            //finally
            //{
                //--ORIG POSITION
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestsService);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequest);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestParams);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oApprovalRequestsParams);
                //GC.Collect();
            //}

        }

        private void SaveDraftToDoc(int draftentry)
        {

            int errmsg = 0;

            //ORIG POSITION
            //SAPbobsCOM.Users oUsers = (SAPbobsCOM.Users)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
            //SAPbobsCOM.Documents oDraft = (SAPbobsCOM.Documents)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);

            ////SAPbobsCOM.Documents oPr = (SAPbobsCOM.Documents)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest);
            //SAPbobsCOM.Recordset oGetAttachPath = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //SAPbobsCOM.Recordset oGetReportName = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //SAPbobsCOM.Recordset oGetPr = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {

                //[07MAY2026]
                SAPbobsCOM.Users oUsers = (SAPbobsCOM.Users)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
                SAPbobsCOM.Documents oDraft = (SAPbobsCOM.Documents)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);

                //SAPbobsCOM.Documents oPr = (SAPbobsCOM.Documents)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest);
                SAPbobsCOM.Recordset oGetAttachPath = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oGetReportName = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbobsCOM.Recordset oGetPr = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oDraft.GetByKey(draftentry);
                //oDraft.Comments = "Created by: PR Approver Addon";
                errmsg = oDraft.SaveDraftToDocument();

                if (errmsg != 0)
                {
                    string error_msg = "Error on Purchase Request Draft with Doc. No. " + oDraft.DocNum;
                    error_msg += ". Error message: " + Program.oGlobalCompany.GetLastErrorDescription().ToString();
                    oUsers.GetByKey(oDraft.UserSign);

                    SAPMessages.sendAlertsToUser(error_msg, oUsers.UserCode, draftentry.ToString());
                    Application.SBO_Application.StatusBar.SetText(Program.oGlobalCompany.GetLastErrorDescription().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                }
                else
                {
                    
                    //oPr.GetByKey(int.Parse(Program.oGlobalCompany.GetNewObjectKey()));

                    oGetPr.DoQuery($"select A.DocEntry, A.DocNum, B.BeginStr from OPRQ A left join NNM1 B on A.Series = B.Series where A.DocEntry = '{Program.oGlobalCompany.GetNewObjectKey()}'");
                    oGetAttachPath.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'attachment_path' ");
                    oGetReportName.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'report_name' ");

                    //Email.CreateEmail("Purchase Request " + oDraft.DocNum + " has been approved.");
                    if (oGetAttachPath.RecordCount > 0 && oGetReportName.RecordCount > 0 && oGetPr.RecordCount > 0)
                    {
                        //Attachmentss.GeneratePDF($"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", $"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", oPr.DocEntry.ToString(), oGetReportName.Fields.Item("Name").Value.ToString());
                        //Email.CreateEmail("Purchase Request " + oPr.SeriesString + "-" + oPr.DocNum.ToString() + " has been approved.", $"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", "PR" + oPr.DocEntry.ToString() + ".pdf");
                        Attachmentss.GeneratePDF($"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", $"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", oGetPr.Fields.Item("DocEntry").Value.ToString(), oGetReportName.Fields.Item("Name").Value.ToString());
                        Email.CreateEmail("Purchase Request " + oGetPr.Fields.Item("BeginStr").Value.ToString() + "-" + oGetPr.Fields.Item("DocNum").Value.ToString() + " has been approved.", $"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", "PR" + oGetPr.Fields.Item("DocEntry").Value.ToString() + ".pdf");

                    }
                    else
                    {
                        //Application.SBO_Application.StatusBar.SetText("Email will be sent for " + "Purchase Request " + oPr.DocNum.ToString() + " without attachment. No attachment path defined.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        Application.SBO_Application.StatusBar.SetText("Email will be sent for " + "Purchase Request " + oGetPr.Fields.Item("DocNum").Value.ToString() + " without attachment. No attachment path defined.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }


                }

                //[07MAY2026]
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUsers);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDraft);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oPr);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetAttachPath);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetReportName);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetPr);
                GC.Collect();
            }
            catch (Exception ex)
            {
                Logger.WriteToFile("ERROR", "SaveDraftToDoc", ex.Message);
                Application.SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            //finally
            //{   
                //ORIG POSITION
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oUsers);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oDraft);
                ////System.Runtime.InteropServices.Marshal.ReleaseComObject(oPr);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetAttachPath);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetReportName);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetPr);
             //   GC.Collect();
            //}

        }


        private void SaveDraftToDoc_2(int draftentry)
        {

            //[20MAY2026] Created to Improve the SAVEDRAFT
            try
            {
                SAPbobsCOM.Users oUsers = (SAPbobsCOM.Users)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
                SAPbobsCOM.Documents oDraft = (SAPbobsCOM.Documents)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);

                SAPbobsCOM.Recordset oGetAttachPath = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oGetReportName = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oGetPr = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // 1. Verify the Draft actually exists before trying to save it
                if (oDraft.GetByKey(draftentry))
                {
                    int errmsg = oDraft.SaveDraftToDocument();

                    if (errmsg != 0)
                    {
                        // 2. Properly extract the specific DI API error code AND message
                        int errCode;
                        string errMsgDI;
                        Program.oGlobalCompany.GetLastError(out errCode, out errMsgDI);

                        // Construct a detailed error string
                        string error_msg = $"Error on PR Draft DocNum {oDraft.DocNum}. DI Error [{errCode}]: {errMsgDI}";

                        oUsers.GetByKey(oDraft.UserSign);
                        SAPMessages.sendAlertsToUser(error_msg, oUsers.UserCode, draftentry.ToString());
                     // SAPMessages.sendAlertsToUser(error_msg, oUsers.UserCode, draftentry.ToString());   //ORIG


                        // 3. Print the detailed string to the Status Bar, and set duration to Long
                        Application.SBO_Application.StatusBar.SetText(error_msg, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                     // Application.SBO_Application.StatusBar.SetText(Program.oGlobalCompany.GetLastErrorDescription().ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);  //ORG

                    }
                    else
                    {
                        //21MAY2026
                        Application.SBO_Application.StatusBar.SetText("Document Draft " + draftentry + " is now approved.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);


                        oGetPr.DoQuery($"select A.DocEntry, A.DocNum, B.BeginStr from OPRQ A left join NNM1 B on A.Series = B.Series where A.DocEntry = '{Program.oGlobalCompany.GetNewObjectKey()}'");
                        oGetAttachPath.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'attachment_path' ");
                        oGetReportName.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'report_name' ");

                        if (oGetAttachPath.RecordCount > 0 && oGetReportName.RecordCount > 0 && oGetPr.RecordCount > 0)
                        {
                            
                            Attachmentss.GeneratePDF($"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", $"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", oGetPr.Fields.Item("DocEntry").Value.ToString(), oGetReportName.Fields.Item("Name").Value.ToString());

                            Email.CreateEmail("Purchase Request " + oGetPr.Fields.Item("BeginStr").Value.ToString() + "-" + oGetPr.Fields.Item("DocNum").Value.ToString() + " has been approved.", $"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", "PR" + oGetPr.Fields.Item("DocEntry").Value.ToString() + ".pdf");


                        }
                        else
                        {
                            Application.SBO_Application.StatusBar.SetText("Email will be sent for Purchase Request " + oGetPr.Fields.Item("DocNum").Value.ToString() + " without attachment. No attachment path defined.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }
                }
                else
                {
                    // Handle scenario where draftentry doesn't exist
                    Application.SBO_Application.StatusBar.SetText($"Draft entry {draftentry} could not be found.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }


                //[07MAY2026]
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUsers);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDraft);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oPr);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetAttachPath);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetReportName);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetPr);
                GC.Collect();


            }
            catch (Exception ex)
            {
                // 4. Trap any system-level/C# exceptions (e.g. NullReference, SQL Syntax issues)
                string sysError = $"System Exception: {ex.Message}";
                Application.SBO_Application.StatusBar.SetText(sysError, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }



        }




        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;

        private void Button0_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try  //ADD ERROR LOG [07MAY2026]
            {

                if (string.IsNullOrEmpty(this.EditText0.Value))
                {
                    Application.SBO_Application.StatusBar.SetText("Input SAP User password.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    BubbleEvent = false;
                }

            }          
            catch (Exception ex)
            {
                Logger.WriteToFile("ERROR", "Button0_PressedBefore", ex.Message);   //ADD ERROR LOG [07MAY2026]
                Application.SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); //ADD ERROR LOG [29MAR2026]
            }

        }

        private SAPbouiCOM.Button Button1;

        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            /** for testing purposes
             * 
             * 
             * 
            SAPbobsCOM.Recordset oGetAttachPath = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oGetReportName = (SAPbobsCOM.Recordset)Program.oGlobalCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oGetAttachPath.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'attachment_path' ");
            oGetReportName.DoQuery("select * from [@DYN_EMAIL_CONFIG] where Code = 'report_name' ");
            Attachmentss.GeneratePDF($"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", $"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", "165", oGetReportName.Fields.Item("Name").Value.ToString() );
            Email.CreateEmail("Purchase Request" + "165" + " has been approved.", $"{oGetAttachPath.Fields.Item("Name").Value.ToString()}", "PR" + "165" + ".pdf");
            **/

            LoadGrid();

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {

            //ORIG ---------------
            //BubbleEvent = true;
            //throw new System.NotImplementedException();
            

            BubbleEvent = true;
            try                    //ADD TRY_CATCH [20MAY2026]
            {

                //throw new System.NotImplementedException();      //JUST REMOVE this ACCORDING to GEMINI [20MAY2026]

            }
            catch (Exception ex)
            {
                Logger.WriteToFile("ERROR", "Button0_ClickBefore", ex.Message);
                Application.SBO_Application.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            

        }


    }
}