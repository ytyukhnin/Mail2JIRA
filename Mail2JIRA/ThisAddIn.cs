using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Net;
using System.IO;
using Mail2JIRA.JiraDTO;
using Mail2JIRA.Properties;

namespace Mail2JIRA
{
    public partial class ThisAddIn
    {
        private MessageWrapper messageDetails;
        private Outlook.Inspectors inspectors;
        private Office.CommandBarButton objEmailToolBarButton;
        
        private const string TOOL_BAR_TAG_EMAIL = "M2JEmailToolBar";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(AddButtonToEmailWindow);
            inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(GetMessageDetails);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // TODO delete the button?
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        /// <summary>
        /// Adds a button on email window.
        /// </summary>
        /// <param name="Inspector">Microsoft.Office.Interop.Outlook.Inspector</param>
        private void AddButtonToEmailWindow(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem objMailItem = (Outlook.MailItem)Inspector.CurrentItem;

            if (Inspector.CurrentItem is Outlook.MailItem)
            {
                objMailItem = (Outlook.MailItem)Inspector.CurrentItem;
                // Delete the existing instance, if applicable.
                foreach (Office.CommandBar objCmd in Inspector.CommandBars)
                {
                    if (objCmd.Name == TOOL_BAR_TAG_EMAIL)
                    {
                        objCmd.Delete();
                    }
                }

                Office.CommandBar objCommandBar = Inspector.CommandBars.Add(TOOL_BAR_TAG_EMAIL, Office.MsoBarPosition.msoBarBottom, false, true);
                objEmailToolBarButton = (Office.CommandBarButton)objCommandBar.Controls.Add(Office.MsoControlType.msoControlButton, 1, missing, missing, true);

                objEmailToolBarButton.Caption = "Create a JIRA ticket";
                objEmailToolBarButton.Style = Office.MsoButtonStyle.msoButtonIconAndCaptionBelow;
                objEmailToolBarButton.FaceId = 349; // button faceid code http://www.kebabshopblues.co.uk/2007/01/04/visual-studio-2005-tools-for-office-commandbarbutton-faceid-property/
                objEmailToolBarButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(objEmailToolBarButton_Click);
                objCommandBar.Visible = true;
            }
        }
        /// <summary>
        /// Click handler for the command button.
        /// </summary>
        /// <param name="ctrl">Office.CommandBarButton</param>
        /// <param name="cancel"></param>
        private void objEmailToolBarButton_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            try
            {
               if (messageDetails == null)
                    return;

               JiraService js = new JiraService(new Uri(Settings.Default.JiraApiUrl), Settings.Default.JiraUserName, Settings.Default.JiraPassword);
                CreateIssueResponseObject ro = js.CreateIssue(new JiraDTO.CreateIssueRequestObject()
                {
                    Fields = new JiraDTO.FieldsObject()
                    {
                        Project = new JiraDTO.ProjectObject()
                        {
                            Key = Settings.Default.JiraDestinationProjectKey
                        },
                        Summary = String.Format("Support request on {0} from {1}", messageDetails.CurrentMessage.Subject.Replace("RE:", String.Empty).Trim(), messageDetails.CurrentMessage.To.Trim()),
                        Description = String.Format("Support request from {0} on\n\n{1}", messageDetails.CurrentMessage.To, messageDetails.CurrentMessage.Body.Trim()),
                        IssueType = new JiraDTO.IssueTypeObject()
                        {
                            Name = Settings.Default.JiraDestiantionIssueType
                        },
                        Assignee = new UserObject()
                        {
                            Name = Settings.Default.JiraUserName
                        }
                    }
                });

                if(ro != null)
                {
                    messageDetails.CurrentMessage.HTMLBody =
                        String.Format("Hello {0},<br/><br/><br/><br/>JIRA issue: {1}<br/>JIRA issue link: <a href='{2}'>{2}</a><br/><br/>{3}",
                        messageDetails.CurrentMessage.To.Trim(), ro.Key, String.Concat(Settings.Default.JiraIssueBrowseUrl, ro.Key), messageDetails.CurrentMessage.HTMLBody);
                }

            }
            catch (WebException ex)
            {
                System.Windows.Forms.MessageBox.Show(String.Format("Error {0}\n{1}", ex.Message, ex.Response != null ? new StreamReader(ex.Response.GetResponseStream()).ReadToEnd() : String.Empty), 
                    "Error Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(String.Format("Error {0}", ex.Message),
                    "Error Message", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Gets the message details for the current email.
        /// </summary>
        /// <param name="Inspector">Microsoft.Office.Interop.Outlook.Inspector</param>
        private void  GetMessageDetails(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            messageDetails = null;
            if (Inspector.CurrentItem is Outlook.MailItem)
            {
                messageDetails = new MessageWrapper((Outlook.MailItem)Inspector.CurrentItem);
            }
        }
    }
}
