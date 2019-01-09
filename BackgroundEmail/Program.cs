using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Net;
using System.Diagnostics;
using System.Net.Mail;
using Newtonsoft.Json;
using System.Data;
using OpenPop.Pop3;
using System.IO;
using OpenPop.Mime;
using System.Text.RegularExpressions;

namespace BackgroundEmail
{
    class Parameters
    {
        // subjects sirve para saber si el correo puede contener la información de los parámetros
        //Body sirve para saber a quienes va dirigido el correo
        //Entity para buscar los montos de los pedidos (denominados "Entity") y determinar si debe ir dirigido a alguien mas


        public string[] Subjects { get; set; } //Palabras para buscar en el titulo del correo
        public string[] Body { get; set; }     //Palabras para buscar en el cuerpo del correo
        public string[] Entity { get; set; }   //Palabras para localizar productos
    }

    class Autonomous
    {
        Parameters Parameters;
        FileManager FileManager;
        Mail Mail;

        List<Message> Messages;
        DataTable Table;

        List<string> Destinatarys;
        List<string> EntityMounts;
        List<string> MessagesBodys;

        public Autonomous()
        {
            Destinatarys = new List<string>();
            EntityMounts = new List<string>();
            MessagesBodys = new List<string>();
            FileManager = new FileManager();
            Mail = new Mail();

        }

        public void Run()
        {
            //while (true)
            //{
            try
            {
                ExecuteProcess();
                Analize();
                //Thread.Sleep(30000); // 5 minutos
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e);
            }
            //}
        }

        public void ExecuteProcess()
        {
            string serializeObject = FileManager.ReadFile(FileManager.Parameters);
            Parameters = StringJsonManager.DeserializeObject(serializeObject);

            Messages = Mail.ReadMails(Messages);

            FileManager.WriteFile(FileManager.Mails, StringJsonManager.SerializeObject(Messages));
            StringBuilder Builder = new StringBuilder();

            //por cada mensaje en la lista de mensajes
            foreach (Message Message in Messages)
            {
                //por cada parámetro de titulo
                foreach (string param1 in Parameters.Subjects)
                {
                    //busca si el mensaje coincide con el parámetro del "titulo del correo"
                    if (Message.Headers.Subject.Contains(param1))
                    {
                        //convierte el cuerpo del mensaje en string

                        OpenPop.Mime.MessagePart plainText = Message.FindFirstHtmlVersion();
                        if (plainText != null)
                        {
                            Builder.Append(plainText.GetBodyAsText());
                        }
                        else
                        {
                            OpenPop.Mime.MessagePart html = Message.FindFirstHtmlVersion();
                            if (html != null)
                            {
                                // We found some html!
                                Builder.Append(html.GetBodyAsText());
                            }
                        }

                        string bodytext = Builder.ToString();
                        MessagesBodys.Add(bodytext);
                        //por cada parámetro de cuerpo
                        foreach (string param2 in Parameters.Body)
                        {
                            //si contiene el cuerpo del texto el parámetro
                            if (bodytext.Contains(param2))
                            {
                                Destinatarys.Add(param2);
                            }
                        }

                        //por cada parámetro de entidad
                        foreach (string param3 in Parameters.Entity)
                        {
                            //si contiene esa entidad en el cuerpo
                            if (bodytext.Contains(param3))
                            {
                                //index -> Obtiene la posición donde se encuentra el parámetro
                                //Count y while -> le resta a index 20 posiciones para obtener los 20
                                //caracteres anteriores al parámetro, para buscar la cantidad en expresiones
                                //de tipo: " quiero 500 (parámetro)". 
                                int index = bodytext.IndexOf(param3);
                                int Count = 20;
                                while (index >= 0 && Count != 0) //se detiene si no hay mas caracteres en el texto antes del parámetro
                                {
                                    index--;
                                    Count--;
                                }

                                string After = bodytext.Substring(index, 20 - Count); //20 caracteres anteriores al parámetro
                                string tempmount = "";
                                for (int i = 0; i < After.Length; i++)
                                {
                                    char digit = Convert.ToChar(After.Substring(1, i));
                                    if (Char.IsDigit(digit))
                                    {
                                        tempmount += digit;
                                    }
                                }
                                if (tempmount.Length > 0)
                                {
                                    EntityMounts.Add("Monto no encontrado");
                                }
                                else
                                {
                                    EntityMounts.Add(tempmount);
                                }
                            }
                        }

                    }
                }
            }
        }

        public void Analize()
        {
            List<string> DestinatarysTemp = new List<string>();
            foreach (string var in Destinatarys)
            {
                DestinatarysTemp.Add(var);
            }

            if (DestinatarysTemp.Count > 0)
            {
                foreach (string dept in DestinatarysTemp)
                {
                    switch (dept)
                    {
                        case "Departamento de Ventas":
                            Destinatarys.Remove(dept);
                            Destinatarys.Add("Sales@ChildMarket.com");
                            break;
                        case "Exteriores":
                            Destinatarys.Remove(dept);
                            Destinatarys.Add("Ext@ChildMarket.com");
                            break;
                        default:
                            Destinatarys.Remove(dept);
                            break;
                    }
                }
            }

            foreach (string mount in EntityMounts)
            {
                int Value = Convert.ToInt32(mount);
                if (Value >= 100)
                    Destinatarys.Add("Boss@ChildMarket.com");
            }


            Console.WriteLine("Proceso de autónomo terminado. Correos que se van a enviar: ");
            for (int i = 0; i < Destinatarys.Count; i++)
            {
                Console.WriteLine("Correo " + (i + 1) + " para:" + Destinatarys[i]);
            }
            for (int i = 0; i < MessagesBodys.Count; i++)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("MENSAJE " + (i + 1) + " :");
                Console.ResetColor();
                Console.WriteLine(MessagesBodys[i]);
            }
        }

        public void CreateParameters()
        {
            Parameters Params = new Parameters();
            Params.Subjects = new string[2];
            Params.Body = new string[1];
            Params.Entity = new string[1];

            Params.Subjects[0] = "Producto";
            Params.Subjects[1] = "Compra";
            Params.Body[0] = "Departamento de Ventas";
            Params.Entity[0] = "Niños";

            FileManager.WriteFile(FileManager.Parameters, StringJsonManager.SerializeObject(Params));
        }
    }

    class FileManager
    {
        public const string HotmailClient = @"c:\HotmailClient";
        public const string Mails = @"c:\HotmailClient\Mails";
        public const string Parameters = @"c:\HotmailClient\Parameters";
       
        public FileManager()
        {
            try
            {
                System.IO.Directory.CreateDirectory(HotmailClient);
                Console.WriteLine("Carpeta c:\\HotmailClient creada");
                System.IO.Directory.CreateDirectory(Mails);
                Console.WriteLine("Carpeta c:\\HotmailClient\\Mails creada");
                System.IO.Directory.CreateDirectory(Parameters);
                Console.WriteLine("Carpeta c:\\HotmailClient\\Parameter creada");
            }
            catch (Exception e)
            {
                Console.WriteLine("El proceso de FileManager ha fallado: {0}", e.Message);
            }
        }

        public void WriteFile(string FolderPath, string SerializedObject)
        {
            switch (FolderPath)
            {
                case HotmailClient:
                    FolderPath += @"\HotmailClient.txt";
                    break;
                case Mails:
                    FolderPath += @"\Mails-" + DateTime.Today.Year + DateTime.Today.Month + DateTime.Today.Day + " .txt";
                    break;
                case Parameters:
                    FolderPath += @"\Parameters.txt";
                    break;
            }

            File.WriteAllText(FolderPath, SerializedObject);
        }

        public string ReadFile(string FolderPath)
        {
            switch (FolderPath)
            {
                case HotmailClient:
                    FolderPath += @"\HotmailClient.txt";
                    break;
                case Mails:
                    FolderPath += @"\Mails.txt";
                    break;
                case Parameters:
                    FolderPath += @"\Parameters.txt";
                    break;
            }

            return File.ReadAllText(FolderPath);
        }

    }

    class Mail
    {
        SmtpClient SmtpServer;
        MailMessage MailMessage;
        Pop3Client Pop3Client;

        //Correo y password del correo del autónomo
        string UserMail = "IngEduardoHidalgo@hotmail.com";
        string UserPassword = "23DG/J4E56KTA";

        public void SendMail(string From, List<string> To, string Subject, string Message)
        {
            Console.WriteLine("Mensaje de:   " + From);
            for (int i = 0; i < To.Count; i++)
                Console.WriteLine("Mensaje para: " + To[i]);
            Console.WriteLine("Asunto:       " + Subject);
            Console.WriteLine("Cuerpo:       " + Message);

            try
            {
                using (SmtpServer = new SmtpClient())
                {
                    MailMessage = new MailMessage(From, To[0]);
                    if (To.Count > 1)
                        for (int i = 1; i < To.Count - 1; i++)
                            MailMessage.To.Add(new MailAddress(To[i]));

                    MailMessage.Subject = Subject;
                    MailMessage.Body = Message;

                    SmtpServer.Port = 587;
                    SmtpServer.Host = "smtp-mail.outlook.com";
                    SmtpServer.UseDefaultCredentials = false;
                    SmtpServer.Credentials = new System.Net.NetworkCredential(UserMail, UserPassword);
                    SmtpServer.EnableSsl = true;
                    SmtpServer.Send(MailMessage);
                }
                Console.WriteLine("Mensaje enviado exitosamente");
            }
            catch (Exception e)
            {
                Console.WriteLine("El proceso de SendMail ha fallado: {0}", e.Message);
            }
        }

        public List<Message> ReadMails(List<Message> allMessages)
        {
            try
            {
                using (Pop3Client = new Pop3Client())
                {
                    Console.WriteLine("Obteniendo los mensajes recientes de la bandeja...");
                    Pop3Client.Connect("pop3.live.com", 995, true);
                    Pop3Client.Authenticate(UserMail, UserPassword);

                    int messageCount = Pop3Client.GetMessageCount();
                    Console.WriteLine("Mensajes obtenidos: " + messageCount);
                    allMessages = new List<Message>(messageCount);
                    for (int i = 1; i <= messageCount; i++)
                        allMessages.Add(Pop3Client.GetMessage(i));
                    return allMessages;
                }
                Console.WriteLine("Mensajes obtenidos exitosamente");
            }
            catch (Exception e)
            {
                Console.WriteLine("El proceso de ReadMail ha fallado: {0}", e.Message);
                return allMessages;
            }

        }
    }

    static class StringJsonManager
    {
        public static string SerializeObject(object Obj)
        {
            return JsonConvert.SerializeObject(Obj);
        }

        public static Parameters DeserializeObject(string serialize)
        {
            return JsonConvert.DeserializeObject<Parameters>(serialize);
        }
    }

    class Program
    {

        static void Main(string[] args)
        {
            Autonomous A = new Autonomous();
            A.Run();
            Console.ReadKey();
        }
    }
}
