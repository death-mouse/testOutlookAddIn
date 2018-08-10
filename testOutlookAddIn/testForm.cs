using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace testOutlookAddIn
{
    public partial class testForm : Form
    {
        public testForm()
        {
            InitializeComponent();

        }
        Outlook.MailItem mailItem;
        string analiticEmail;
        private BindingList<TaskCategory> taskCategoriesList= new BindingList<TaskCategory>();
        private BindingList<TaskBU> taskBUList = new BindingList<TaskBU>();

        /// <summary>
        /// Передача на форму текущего сообщения
        /// </summary>
        /// <param name="_mailItem">Текущее сообщение</param>
        public void parmMessage(Outlook.MailItem _mailItem)
        {
            mailItem = _mailItem;
        }
        /// <summary>
        /// Установить почту аналитика
        /// </summary>
        /// <param name="_analiticEmail">Почта аналитика</param>
        public void setAnaliticEmail(string _analiticEmail)
        {
            analiticEmail = _analiticEmail;
        }

        private void testForm_Load(object sender, EventArgs e)
        {
            initComoBoxes();
            const string PR_SMTP_ADDRESS =
        "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            txtMailBody.Text = mailItem.Body;
            txtMailSybject.Text = mailItem.Subject;
            foreach (Outlook.Recipient recipeint in mailItem.Recipients)
            {
                Outlook.PropertyAccessor pa = recipeint.PropertyAccessor;
                string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                txtRecipient.Text += smtpAddress + ";";
            }
            DateTime creationTime = mailItem.CreationTime;
            txtDateTimeCreated.Text = creationTime.ToString(@"yyyy-MM-ddTHH:mm:ss");
            Outlook.PropertyAccessor pa2 =  mailItem.Sender.PropertyAccessor;
            string smtpAddressAuthor = pa2.GetProperty(PR_SMTP_ADDRESS).ToString();
            txtAuthor.Text = smtpAddressAuthor;
        }
        /// <summary>
        /// Инициализация списокв БЕ и категорий
        /// </summary>
        private void initComoBoxes()

        {
            try
            {
                var url = "http://zskpk02:8280/services/sdapi/CategoryList";
                HttpWebRequest request =
                (HttpWebRequest)WebRequest.Create(url);

                request.Method = "GET";
                request.UserAgent = "AddIns by DeAmouSE";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                StringBuilder output = new StringBuilder();
                output.Append(reader.ReadToEnd());
                response.Close();
                string xml = output.ToString();
                xml = xml.Replace(" xmlns=\"sd\"", "");
                XDocument xdoc = XDocument.Parse(xml);
                foreach (XElement taskCategoryElement in xdoc.Element("TaskCategoryCollection").Elements("TaskCategory"))
                {
                    //cmbCategory.Items.Add(new TaskCategory(taskCategoryElement.Element("categoryName").Value, Convert.ToInt16(taskCategoryElement.Element("Id").Value)));
                    taskCategoriesList.Add(new TaskCategory(taskCategoryElement.Element("categoryName").Value, Convert.ToInt16(taskCategoryElement.Element("Id").Value)));
                }
                cmbCategory.DataSource = taskCategoriesList;
                cmbCategory.DisplayMember = "categoryName";
                cmbCategory.ValueMember = "Id";

                url = "http://zskpk02:8280/services/sdapi/BUList";
                request = (HttpWebRequest)WebRequest.Create(url);

                request.Method = "GET";
                request.UserAgent = "AddIns by DeAmouSE";
                response = (HttpWebResponse)request.GetResponse();
                reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                output = new StringBuilder();
                output.Append(reader.ReadToEnd());
                response.Close();
                xml = output.ToString();
                xml = xml.Replace(" xmlns=\"sd\"", "");
                xdoc = XDocument.Parse(xml);

                foreach (XElement taskCategoryElement in xdoc.Element("TaskBUList").Elements("TaskBU"))
                {
                    //cmbBU.Items.Add(new TaskCategory(taskCategoryElement.Element("BuName").Value, Convert.ToInt16(taskCategoryElement.Element("Id").Value)));
                    taskBUList.Add(new TaskBU(taskCategoryElement.Element("BuName").Value, Convert.ToInt16(taskCategoryElement.Element("Id").Value)));
                }

                cmbBU.DataSource = taskBUList;
                cmbBU.DisplayMember = "buName";
                cmbBU.ValueMember = "Id";
            }
            catch(Exception e)
            {
                throw new Exception(e.Message);
            }

        }

        /// <summary>
        /// Класс категорий для добавления в комбобокс
        /// </summary>
        public class TaskCategory
        {
            public TaskCategory(string _categoryName, int _id)
            {
                this.categoryName = _categoryName; this.Id = _id;
            }
            public string categoryName{ get; set; }
            public int Id{ get; set; }
        }
        /// <summary>
        /// Класс БЕ для добавления в комбобокс
        /// </summary>
        public class TaskBU
        {
            public TaskBU(string _buName, int _id)
            {
                this.buName = _buName; this.Id = _id;
            }
            public string buName { get; set; }
            public int Id { get; set; }
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            string xmlDataSend = this.getXmlToSend();
            postXMLData("http://zskpk02:8280/sdapi/task", xmlDataSend);
            this.Close();
        }

        /// <summary>
        /// Формирования xml для отправления его на сервере
        /// </summary>
        /// <returns>xml для отправки на сервер</returns>
        public string getXmlToSend()
        {
            XDocument doc = new XDocument(new XElement("TaskRequest",
                                                new XElement("Author", txtAuthor.Text),
                                                new XElement("Analitic", analiticEmail),
                                                new XElement("TaskDate", txtDateTimeCreated.Text),
                                                new XElement("BUId", ((TaskBU)cmbBU.SelectedItem).Id),
                                                new XElement("CategoryId", ((TaskCategory)cmbCategory.SelectedItem).Id),
                                                new XElement("Subject", txtMailSybject.Text),
                                                new XElement("Body", txtMailBody.Text)));
            return doc.ToString();
        }

        /// <summary>
        /// Отправка данных на веб сервис для создания заявки.
        /// </summary>
        /// <param name="destinationUrl">Ссылка на веб сервис</param>
        /// <param name="requestXml">XML данные</param>
        public void postXMLData(string destinationUrl, string requestXml)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(destinationUrl);
                byte[] bytes;
                bytes = System.Text.Encoding.UTF8.GetBytes(requestXml);
                request.ContentType = "text/xml; encoding='utf-8'";
                request.ContentLength = bytes.Length;
                request.Method = "POST";
                request.UserAgent = "AddIns by DeAmouSE";
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(bytes, 0, bytes.Length);
                requestStream.Close();
                HttpWebResponse response;
                response = (HttpWebResponse)request.GetResponse();
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream responseStream = response.GetResponseStream();
                    string responseStr = new StreamReader(responseStream).ReadToEnd();
                    XDocument xdoc = XDocument.Parse(responseStr);
                    if (xdoc.Element("result").Element("error") != null)
                    {
                        string errorText = xdoc.Element("result").Element("error").Value;
                        if (errorText != "")
                        {
                            MessageBox.Show(errorText, "Ошибка при создании заявки", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                            return;
                        }
                    }
                    string taskId = xdoc.Element("result").Element("taskId").Value;
                    if (taskId != "")
                    {
                        MessageBox.Show(string.Format("Создана заявка {0}. Ссылка на заявку скопирована в буфер обмена", taskId), "Заявка успешно создана", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                        Clipboard.SetText(string.Format("http://servicedesk.gradient.ru/Task/View/{0}", taskId));
                    }
                    else
                    {
                        MessageBox.Show("Странная ошибка, очень странная...", "Ошибка при создании заявки", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
            }
            catch ( Exception e)
            {
                MessageBox.Show(e.Message, "Ошибка при создании заявки", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }

        }
    }
}
