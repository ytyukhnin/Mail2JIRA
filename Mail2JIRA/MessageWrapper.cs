using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Mail2JIRA
{
    /// <summary>
    /// Message wrapper.
    /// </summary>
    public class MessageWrapper
    {
        private Microsoft.Office.Interop.Outlook.MailItem currentMessage;

        public Microsoft.Office.Interop.Outlook.MailItem CurrentMessage { get { return currentMessage; } }

        public MessageWrapper(Microsoft.Office.Interop.Outlook.MailItem currentMessage)
        {
            this.currentMessage = currentMessage;
        }
    }
}
