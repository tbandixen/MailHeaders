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
    public static class StringExtentions
    {
        public static string NullIfEmpty(this string str)
        {
            return string.IsNullOrWhiteSpace(str) ? null : str;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Username");
            var username = Console.ReadLine().NullIfEmpty();
            Console.WriteLine("Password");
            var password = Console.ReadLine().NullIfEmpty();
            Console.WriteLine("Host");
            var host = Console.ReadLine().NullIfEmpty();
            Console.WriteLine("UseSSL [Y/n]");
            var useSsl = (Console.ReadLine().ToLower().NullIfEmpty() ?? "y") == "y";
            Console.WriteLine("Port (110/995)");
            var port = Convert.ToInt32(Console.ReadLine().NullIfEmpty() ?? (useSsl ? "995" : "110"));
            Console.WriteLine("Amount (0 for all)");
            var amount = Convert.ToInt32(Console.ReadLine().NullIfEmpty() ?? "0");
            Console.WriteLine("Start from (0 for the beginning)");
            var offset = Convert.ToInt32(Console.ReadLine().NullIfEmpty() ?? "0");

            Console.Clear();

            Console.WriteLine("Getting the messages from POP Server...");

            string filename = $"{DateTime.Now.ToLocalTime().ToString("yyMMddHHmmss")}-Analyse-{host}.csv";
#if DEBUG
            filename = $"_{Path.GetFileNameWithoutExtension(Path.GetRandomFileName())}.csv";

            if (amount == 0)
            {
                amount = 5;
            }
#endif
            using (FileStream filestream = new FileStream(filename, FileMode.Create))
            using (var streamwriter = new StreamWriter(filestream) { AutoFlush = true })
            {
                streamwriter.WriteLine($"\"{"Id"}\";\"{"Date sent"}\";\"{"Received date"}\";\"{"Delay"}\";\"{"Received from"}\";\"{"Received by"}\"");
                try
                {
                    WriteMessages(streamwriter, username, password, host, useSsl, port, amount, offset);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    Console.ReadLine();
                }
            }
        }

        private static void LogMessage(TextWriter writer, MessageHeader header)
        {
            DateTime lastReceivedDate = header.DateSent.ToLocalTime();
            /**
             * https://mediatemple.net/community/products/dv/204643950/understanding-an-email-header
             * The received is the most important part of the email header and is usually the most reliable.
             * They form a list of all the servers/computers through which the message traveled in order to reach you.
             * The received lines are best read from bottom to top. That is, the first "Received:" line is your own system or mail server.
             * The last "Received:" line is where the mail originated. Each mail system has their own style of "Received:" line. A "Received:" line typically identifies the machine that received the mail and the machine from which the mail was received.
             */

            var received = header.Received; //.OrderBy(r => r.Date);
            received.Reverse();
            foreach (var r in received)
            {
                var from = r.Names.Where(n => n.Key == "from").Select(n => n.Value).FirstOrDefault();
                var by = r.Names.Where(n => n.Key == "by").Select(n => n.Value).FirstOrDefault();

                DateTime? receivedDate = null;
                TimeSpan? delay = null;
                if (r.Date != DateTime.MinValue)
                {
                    receivedDate = r.Date.ToLocalTime();
                    delay = receivedDate.Value.Subtract(lastReceivedDate);
                }
                lastReceivedDate = receivedDate.GetValueOrDefault(header.DateSent.ToLocalTime());

                writer.WriteLine($"\"{header.MessageId}\";\"{header.DateSent.ToLocalTime()}\";\"{(receivedDate.HasValue ? receivedDate.ToString() : "")}\";\"{delay}\";\"{from}\";\"{by}\"");
            }
        }

        private static void LogProgress(int current, int total)
        {
            var progress = Math.Round(decimal.Divide(current, total), 1, MidpointRounding.AwayFromZero) * 100;
            Console.Write($"\r{progress}% [{current}/{total}]          ");
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
        public static IEnumerable<MessageHeader> WriteMessages(TextWriter writer, string username, string password, string hostname, bool useSsl = true, int port = 0, int amount = 0, int offset = 0)
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
                int start = messageCount - offset;

                if (amount == 0)
                {
                    amount = start;
                }
                int end = start - amount;

                ICollection<MessageHeader> allMessages = new List<MessageHeader>(amount);

                for (int i = start; i > end; i--)
                {
                    MessageHeader msg;
                    try
                    {
                        msg = client.GetMessageHeaders(i);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine($"Failed at message nr \"{i}\"");
                        throw;
                    }
                    allMessages.Add(msg);
                    LogMessage(writer, msg);
                    LogProgress(start - i + 1, start - end);
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
