//--------------------------------------------------------------------------------------------------

// IHM Engenharia e Sistemas de Automação Ltda

// Módulo: Salva anexos em um diretório específico

// Cliente: Vale S.A.

// Autor Criação: Marcelo Cunha

// Data Criação: 21/02/2018 11:22:00

// Observações: N/A

//--------------------------------------------------------------------------------------------------

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Limilabs.Client.IMAP;
using Limilabs.Client.POP3;
using Limilabs.Client.SMTP;
using Limilabs.Mail;
using Limilabs.Mail.MIME;
using Limilabs.Mail.Fluent;
using Limilabs.Mail.Headers;
using System.Configuration;
using System.Collections.Specialized;


namespace SalvaAnexos
{

    class Program
    {

        static void Main(string[] args)
        {

            using (Imap imap = new Imap())
            {

                try
                {

                    MimeType tipoArquivo = null;

                    string path;

                    string emailSender;
                    emailSender = ConfigurationManager.AppSettings.Get("EmailSender");

                    string imapServer;
                    imapServer = ConfigurationManager.AppSettings.Get("ImapServer");

                    string emailAccount;
                    emailAccount = ConfigurationManager.AppSettings.Get("EmailAccount");

                    string password;
                    password = ConfigurationManager.AppSettings.Get("Password");

                    string fileType;
                    fileType = ConfigurationManager.AppSettings.Get("FileType");

                    imap.Connect(imapServer); // or ConnectSSL for SSL
                    imap.UseBestLogin(emailAccount, password);
                    imap.SelectInbox();

                    List<long> uidList = imap.Search(Flag.All);
                    foreach (long uid in uidList)
                    {
                        IMail email = new MailBuilder()
                            .CreateFromEml(imap.GetMessageByUID(uid));

                        if (email.Sender.Address == emailSender)
                        {

                            foreach (MimeData mime in email.Attachments)
                            {

                                tipoArquivo = mime.ContentType.MimeType;

                                if (Convert.ToString(tipoArquivo) == fileType)
                                {

                                    path = Convert.ToString(@"c:\" + mime.SafeFileName);

                                    if (!File.Exists(path))
                                    {

                                        mime.Save(@"c:\" + mime.SafeFileName);

                                    }

                                }

                            }

                        }

                    }

                    imap.Close();
                }

                catch (Exception ex)
                {

                }

            }

        }

    }

}