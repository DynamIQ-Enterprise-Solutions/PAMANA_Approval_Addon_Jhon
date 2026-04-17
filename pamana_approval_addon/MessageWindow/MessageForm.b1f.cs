using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace pamana_approval_addon
{
    [FormAttribute("MessageForm", "MessageWindow/MessageForm.b1f")]
    class MessageForm : UserFormBase
    {
        public MessageForm()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("msglbl").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {

        }

        public static void displayStatus(string msg, SAPbouiCOM.Form _oForm)
        {
            SAPbouiCOM.Form form;
            SAPbouiCOM.Item item;
            SAPbouiCOM.FormCreationParams formCreationParam;
            try
            {
                formCreationParam = (SAPbouiCOM.FormCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                formCreationParam.UniqueID = ("FTSTAT");
                formCreationParam.FormType = ("FTSTAT");
                formCreationParam.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle;
                form = Application.SBO_Application.Forms.AddEx(formCreationParam);
                form.Left = (_oForm.Left + (_oForm.Width / 2) - (_oForm.Width / 6));
                form.Width = (300);
                form.Top = (_oForm.Top + (_oForm.Height / 2) - (_oForm.Height / 6));
                form.Height = (70);
                item = form.Items.Add("stStatus", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                item.Left = (6);
                item.Width = (290);
                item.Top = (25);
                item.Height = (14);
                ((SAPbouiCOM.StaticText)item.Specific).Caption = msg;
                item.Enabled = false;
                form.Visible = true;

            }
            catch
            {
                try
                {
                    form = Application.SBO_Application.Forms.GetForm("FTSTAT", -1);
                    form.Visible = true;
                    item = form.Items.Add("stStatus", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    ((SAPbouiCOM.StaticText)item.Specific).Caption = msg;
                }
                catch
                {

                }
            }
            formCreationParam = null;
            GC.Collect();
        }

        public static void changeStatus(string as_status)
        {
            try
            {
                SAPbouiCOM.Form form = Application.SBO_Application.Forms.GetForm("FTSTAT", -1);
                form.Visible = true;
                SAPbouiCOM.StaticText specific = (SAPbouiCOM.StaticText)form.Items.Item("stStatus").Specific;
                specific.Caption = as_status;
            }
            catch (Exception exception)
            {
            }
        }

        public static void hideStatus()
        {
            try
            {
                //oApp.Forms.Item("FTSTAT").Visible = false;
                Application.SBO_Application.Forms.Item("FTSTAT").Close();
            }
            catch
            {
            }
        }
    }
}
