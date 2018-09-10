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
            //Проверяем, что выбрали сообщение, и делаем его текущем для передачи на форму
            if ((isMailItem.OutlookItem is Outlook.MailItem))
            {
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
            Outlook.Application outlookApplication = this.OutlookFormRegion.Application; //На всякий случай получение приложения Outlook 
            //OlWindowState outlookApplicationWindowState = outlookApplication.ActiveExplorer().WindowState; в каком положении находится окно
            testForm testForm1 = new testForm();
            testForm1.parmMessage(mailItem); //Передаем текущее письмо в форму
            Outlook.Accounts accounts = (Outlook.Accounts)this.OutlookFormRegion.Session.Accounts; //Получение текущего пользователя, под кем запущен OutLook
            testForm1.setAnaliticEmail(accounts[1].SmtpAddress);//Передаем email в форму
            testForm1.ShowDialog();//Отображаем как диалог

            outlookApplication.ActiveExplorer().Activate(); //Возвращает фокус приложение, в противном случае, outlook уходит на задний план после закрытия формы testform1
            //outlookApplication.ActiveExplorer().WindowState = outlookApplicationWindowState; //Вдруг когда то понадобится менять положения окна


        }

        private void btnParameters_Click(object sender, EventArgs e)
        {
            Outlook.Application outlookApplication = this.OutlookFormRegion.Application;
            Outlook.Categories categories = outlookApplication.Session.Categories;
            List<string> categoryList = new List<string>();
            foreach (Outlook.Category category in categories)
            {
                categoryList.Add(category.Name);
            }
            ParametersForm parametersForm = new ParametersForm();
            parametersForm.parmListCategory(categoryList);
            parametersForm.ShowDialog();
            outlookApplication.ActiveExplorer().Activate();

        }
    }
}
