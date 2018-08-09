using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace testOutlookAddIn
{
    partial class FormRegion1
    {
        #region Фабрика областей формы 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("testOutlookAddIn.FormRegion1")]
        public partial class FormRegion1Factory
        {
            // Возникает перед инициализацией области формы.
            // Чтобы исключить появление области формы, задайте для параметра e.Cancel значение true.
            // Используйте e.OutlookItem для получения ссылки на текущий элемент Outlook.
            private void FormRegion1Factory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }

        #endregion

        // Возникает перед отображением области формы.
        // Используйте this.OutlookItem для получения ссылки на текущий элемент Outlook.
        // Используйте this.OutlookFormRegion для получения ссылки на область формы.
        public Outlook.MailItem mailItem;
        private void FormRegion1_FormRegionShowing(object sender, System.EventArgs e)
        {
            var isMailItem = (Microsoft.Office.Tools.Outlook.FormRegionControl)sender;
            if ((isMailItem.OutlookItem is Outlook.MailItem))
            {
                var test = isMailItem.OutlookItem as Outlook.MailItem;
                mailItem = isMailItem.OutlookItem as Outlook.MailItem;
            }
        }
        // Возникает перед закрытием области формы.
        // Используйте this.OutlookItem для получения ссылки на текущий элемент Outlook.
        // Используйте this.OutlookFormRegion для получения ссылки на область формы.
        private void FormRegion1_FormRegionClosed(object sender, System.EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Office.IRibbonControl ctl = sender as Inspector;
            testForm testForm1 = new testForm();
            testForm1.parmMessage(mailItem);
            //MailItem item = sender as MailItem;
            Outlook.Accounts accounts = (Outlook.Accounts)this.OutlookFormRegion.Session.Accounts;
            testForm1.setAnaliticEmail(accounts[1].SmtpAddress);
            testForm1.ShowDialog();
        }
    }
}
