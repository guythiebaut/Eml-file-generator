using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;

namespace GenerateEmlFile
{
    //https://www.codeguru.com/csharp/saving-an-e-mail-as-an-eml-file-in-c/
    internal class EmlFileGenerator
    {
        internal class EmailArgs
        {
            internal string From = string.Empty;
            internal string To = string.Empty;
            internal string Subject = string.Empty;
            internal string Body = string.Empty;
            internal List<string> Files = new List<string>();
            internal string OutputDir = string.Empty;
            internal bool HelpInvoked = false;
        }

        public void CreateEmlFile(string[] args)
        {

            var emailArgs = UnpackArgs(args);

            if (!ArgsAreValid(emailArgs))
            {
                return;
            }

            if (emailArgs.HelpInvoked)
            {
                return;
            }

            using (var client = new SmtpClient())
            {
                var msg = new MailMessage(emailArgs.From, emailArgs.To, emailArgs.Subject, emailArgs.Body);
                msg.Attachments.Clear();

                foreach (var file in emailArgs.Files)
                {
                    msg.Attachments.Add(new Attachment(file));
                }

                client.UseDefaultCredentials = true;
                client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                client.PickupDirectoryLocation = emailArgs.OutputDir;

                try
                {
                    client.Send(msg);
                    Console.WriteLine();
                    Console.WriteLine($".eml file written to {emailArgs.OutputDir}.");
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception caught: {0}", ex.ToString());
                    Console.ReadLine();
                    Environment.Exit(-1);
                }
            }
        }

        private EmailArgs UnpackArgs(string[] args)
        {
            var emailArgs = new EmailArgs();

            if (!args.Any())
            {
                emailArgs.HelpInvoked = true;
                DisplayHelp();
                return emailArgs;
            }

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-d":
                        emailArgs.OutputDir = args[i + 1];
                        i++;
                        break;
                    case "-f":
                        emailArgs.From = args[i + 1];
                        i++;
                        break;
                    case "-t":
                        emailArgs.To = args[i + 1];
                        i++;
                        break;
                    case "-s":
                        emailArgs.Subject = args[i + 1];
                        i++;
                        break;
                    case "-b":
                        emailArgs.Body = args[i + 1];
                        i++;
                        break;
                    case "-a":
                        emailArgs.Files = UnpackFiles(args, i + 1);
                        i += emailArgs.Files.Count();
                        break;
                    case "-h":
                        emailArgs.HelpInvoked = true;
                        DisplayHelp();
                        i++;
                        break;
                    default:
                        emailArgs.HelpInvoked = true;
                        DisplayHelp();
                        i++;
                        break;
                }

                if (emailArgs.HelpInvoked)
                {
                    break;
                }
            }

            return emailArgs;
        }

        private void DisplayHelp()
        {
            Console.WriteLine("");
            Console.WriteLine("-d destination folder where eml file is written.");
            Console.WriteLine("-f from email address.");
            Console.WriteLine("-t to email address.");
            Console.WriteLine("-s subject text.");
            Console.WriteLine("-b body text.");
            Console.WriteLine("-h help.");
            Console.WriteLine("-a attachments - quoted attachment paths separated by a space.");
            Console.WriteLine("");
            Console.WriteLine("Example");
            Console.WriteLine("");
            Console.WriteLine("-d \"c:\\temp\" -f \"mail@gmail.com\" -t \"import@mail.com\" -s \"Test email\" -b \"Test email body.\" -a \"c:\\temp\\test.txt\" \"c:\\temp\\test2.txt\"");
            Console.WriteLine("");
        }

        private List<string> UnpackFiles(string[] args, int position)
        {
            var files = new List<string>();
            var otherSwitches = new List<string> { "-d", "-f", "-t", "-s", "-b", "-h" };

            for (int i = position; i < args.Length; i++)
            {
                if (otherSwitches.Contains(args[i]))
                {
                    break;
                }
                files.Add(args[i]);
            }

            return files;
        }

        private bool ArgsAreValid(EmailArgs emailArgs)
        {
            bool validArgs = true;

            if (emailArgs.From.Trim() == string.Empty || !IsEmailAddress(emailArgs.From))
            {
                Console.WriteLine("-f argument needs to contain a valid email address.");
                validArgs = false;
            }

            if (emailArgs.To.Trim() == string.Empty || !IsEmailAddress(emailArgs.To))
            {
                Console.WriteLine("-t argument needs to contain a valid email address.");
                validArgs = false;
            }

            if (emailArgs.OutputDir.Trim() == string.Empty || !Directory.Exists(emailArgs.OutputDir))
            {
                Console.WriteLine("-d argument needs to contain a valid destination path for the eml file.");
                validArgs = false;
            }

            return validArgs;
        }

        private bool IsEmailAddress(string email)
        {
            string emailPattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            return Regex.IsMatch(email, emailPattern);
        }
    }
}

