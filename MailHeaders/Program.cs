using OpenPop.Mime.Header;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace MailHeaders
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Username");
            var username = Console.ReadLine();
            Console.WriteLine("Password");
            var password = Console.ReadLine();
            Console.WriteLine("Host");
            var host = Console.ReadLine();
            Console.WriteLine("UseSSL [y/n]");
            var useSsl = Console.ReadLine().ToLower() == "y";
            Console.WriteLine("Port (110/995)");
            var port = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("Amount (0 for all)");
            var amount = Convert.ToInt32(Console.ReadLine());

            FileStream filestream = new FileStream(DateTime.Now.ToString("yyMMddHHmmss") + "-Analyse-" + host + ".csv", FileMode.Create);
            var streamwriter = new StreamWriter(filestream);
            streamwriter.AutoFlush = true;
            var output = Console.Out;
            Console.SetOut(streamwriter);

            var list = FetchAllMessages(username, password, host, useSsl, port, amount);

            Console.WriteLine("\"{0}\";\"{1}\";\"{2}\";\"{3}\";\"{4}\";\"{5}\"", "Date sent", "Subject", "Received date", "Delay", "Received by", "Received from");
            list.ForEach(LogMessage);
        }

        private static void LogMessage(MessageHeader header)
        {
            DateTime lastDate = header.DateSent;
            header.Received.OrderBy(r => r.Date).ToList().ForEach((r) =>
            {
                TimeSpan delay = r.Date - lastDate;
                if (delay < TimeSpan.FromSeconds(0))
                {
                    delay = TimeSpan.FromSeconds(0);
                }
                lastDate = r.Date;
                var from = r.Names.Where(n => n.Key == "from").Select(n => n.Value).FirstOrDefault();
                var by = r.Names.Where(n => n.Key == "by").Select(n => n.Value).FirstOrDefault();
                WL((by ?? "").StartsWith("mail.vs-online.net") ? ConsoleColor.Red : ConsoleColor.White, "\"{0}\";\"{1}\";\"{2}\";\"{3}\";\"{4}\";\"{5}\"", header.DateSent, header.Subject, r.Date, delay, by, from);
            });
        }

        private static void WL(ConsoleColor color, string format, params object[] args)
        {
            var oldColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(format, args);
            Console.ForegroundColor = oldColor;
        }

        /// <summary>
        /// Example showing:
        ///  - how to fetch all messages from a POP3 server
        /// </summary>
        /// <param name="hostname">Hostname of the server. For example: pop3.live.com</param>
        /// <param name="port">Host port to connect to. Normally: 110 for plain POP3, 995 for SSL POP3</param>
        /// <param name="useSsl">Whether or not to use SSL to connect to server</param>
        /// <param name="username">Username of the user on the server</param>
        /// <param name="password">Password of the user on the server</param>
        /// <returns>All Messages on the POP3 server</returns>
        public static List<MessageHeader> FetchAllMessages(string username, string password, string hostname, bool useSsl = true, int port = 0, int amount = 0)
        {
            if (port == 0)
            {
                port = useSsl ? 995 : 110;
            }
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl, 60000, 60000, certificateValidator);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Get the number of messages in the inbox
                int messageCount = client.GetMessageCount();

                // We want to download all messages
                List<MessageHeader> allMessages = new List<MessageHeader>(messageCount);

                // Messages are numbered in the interval: [1, messageCount]
                // Ergo: message numbers are 1-based.
                // Most servers give the latest message the highest number

                if (amount == 0)
                {
                    amount = messageCount;
                }

                for (int i = messageCount; i > messageCount - amount; i--)
                {
                    var msg = client.GetMessageHeaders(i); // .GetMessage(i);
                    allMessages.Add(msg);
                }

                // Now return the fetched messages
                return allMessages;
            }
        }
        private static bool certificateValidator(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslpolicyerrors)
        {
            // We should check if there are some SSLPolicyErrors, but here we simply say that
            // the certificate is okay - we trust it.
            return true;
        }
    }
}
