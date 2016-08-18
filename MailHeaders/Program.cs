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
#if DEBUG
                amount = 5;
#endif

            using (FileStream filestream = new FileStream(DateTime.Now.ToLocalTime().ToString("yyMMddHHmmss") + "-Analyse-" + host + ".csv", FileMode.Create))
            using (var streamwriter = new StreamWriter(filestream))
            {
                //streamwriter.AutoFlush = true;
                //var output = Console.Out;
                Console.SetOut(streamwriter);

                var list = FetchAllMessages(username, password, host, useSsl, port, amount);

                Console.WriteLine($"\"{"Id"}\";\"{"Date sent"}\";\"{"Received date"}\";\"{"Delay"}\";\"{"Received by"}\";\"{"Received from"}\"");
                list.OrderByDescending(i => i.DateSent).ToList().ForEach(LogMessage);
            }
        }

        private static void LogMessage(MessageHeader header)
        {
            DateTime lastReceivedDate = header.DateSent.ToLocalTime();
            header.Received.OrderBy(r => r.Date).ToList().ForEach((r) =>
            {
                var receivedDate = r.Date == DateTime.MinValue ? header.DateSent.ToLocalTime() : r.Date.ToLocalTime();
                TimeSpan delay = receivedDate.Subtract(lastReceivedDate);
                lastReceivedDate = receivedDate;
                //if (delay < TimeSpan.FromSeconds(0))
                //{
                //    delay = TimeSpan.FromSeconds(0);
                //}
                var from = r.Names.Where(n => n.Key == "from").Select(n => n.Value).FirstOrDefault();
                var by = r.Names.Where(n => n.Key == "by").Select(n => n.Value).FirstOrDefault();
                Console.WriteLine($"\"{header.MessageId}\";\"{header.DateSent.ToLocalTime()}\";\"{receivedDate}\";\"{delay}\";\"{by}\";\"{from}\"");
            });
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
        public static IEnumerable<MessageHeader> FetchAllMessages(string username, string password, string hostname, bool useSsl = true, int port = 0, int amount = 0)
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

                if (amount == 0)
                {
                    amount = messageCount;
                }

                ICollection<MessageHeader> allMessages = new List<MessageHeader>(amount);

                for (int i = messageCount; i > messageCount - amount; i--)
                {
                    var msg = client.GetMessageHeaders(i);
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
