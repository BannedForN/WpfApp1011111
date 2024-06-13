using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfApp10
{
    public partial class SendWindow : Window
    {
        public SendWindow()
        {
            InitializeComponent();
        }
        private void SendEmailButton_Click(object sender, RoutedEventArgs e)
        {
            // Send the current Excel file via email
            MailMessage mail = new MailMessage(FromTbx.Text, ToTbx.Text, TitleTbx.Text, "Content of the email");

            Attachment attachment = new Attachment(MainWindow.currentFilePath);
            mail.Attachments.Add(attachment);

            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587);
            smtpClient.Credentials = new System.Net.NetworkCredential(FromTbx.Text, PswTbx.Text);
            smtpClient.EnableSsl = true;
            smtpClient.Send(mail);
        }
    }
}
