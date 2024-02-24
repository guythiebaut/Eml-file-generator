using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;
using static GenerateEmlFile.EmlFileGenerator;

namespace GenerateEmlFile
{
    //https://www.codeguru.com/csharp/saving-an-e-mail-as-an-eml-file-in-c/
    internal class EmlFileGenerator
    {
        internal class EmailArgs
        {
            internal string? From;
            internal string? To;
            internal string? Subject;
            internal string? Body;
            internal List<string> Files = new List<string>();
            internal string? OutputDir;
            internal bool HelpInvoked;
        }

        public void CreateEmlFile(string[] args)
        {

            var emailArgs = UnpackArgs(args);

            if (emailArgs.HelpInvoked)
            {
                return;
            }

            if (!ArgsAreValid(emailArgs))
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
                        emailArgs.OutputDir = GetNextArg(args, i);
                        RequestHelpIfNeeded(ref emailArgs, emailArgs.OutputDir);
                        i++;
                        break;
                    case "-f":
                        emailArgs.From = GetNextArg(args, i);
                        RequestHelpIfNeeded(ref emailArgs, emailArgs.From);
                        i++;
                        break;
                    case "-t":
                        emailArgs.To = GetNextArg(args, i);
                        RequestHelpIfNeeded(ref emailArgs, emailArgs.To);
                        i++;
                        break;
                    case "-s":
                        emailArgs.Subject = GetNextArg(args, i);
                        i++;
                        break;
                    case "-b":
                        emailArgs.Body = GetNextArg(args, i);
                        i++;
                        break;
                    case "-a":
                        emailArgs.Files = UnpackFiles(args, i + 1);
                        i += emailArgs.Files.Count();
                        break;
                    case "-h":
                        emailArgs.HelpInvoked = true;
                        i++;
                        break;
                    default:
                        emailArgs.HelpInvoked = true;
                        i++;
                        break;
                }

                if (emailArgs.HelpInvoked)
                {
                    DisplayHelp();
                    break;
                }
            }

            return emailArgs;
        }

        private EmailArgs RequestHelpIfNeeded(ref EmailArgs emailArgs, string value)
        {
            if (value == string.Empty)
            {
                emailArgs.HelpInvoked = true;
            }

            return emailArgs;
        }

        private string GetNextArg(string[] args, int i)
        {
            string returnValue;

            try
            {
                returnValue = args[i + 1];
            }
            catch (Exception)
            {
                returnValue = string.Empty;
            }

            return returnValue;
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
            Console.WriteLine($"{AppDomain.CurrentDomain.FriendlyName} -d \"c:\\temp\" -f \"mail@gmail.com\" -t \"import@mail.com\" -s \"Test email\" -b \"Test email body.\" -a \"c:\\temp\\test.txt\" \"c:\\temp\\test2.txt\"");
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

            if (!EmailValidated(emailArgs.From))
            {
                Console.WriteLine("-f argument needs to contain a valid email address.");
                validArgs = false;
            }

            if (!EmailValidated(emailArgs.To))
            {
                Console.WriteLine("-t argument needs to contain a valid email address.");
                validArgs = false;
            }

            if (string.IsNullOrEmpty(emailArgs.OutputDir) || !Directory.Exists(emailArgs.OutputDir))
            {
                Console.WriteLine("-d argument needs to contain a valid destination path for the eml file.");
                validArgs = false;
            }

            return validArgs;
        }

        private bool EmailValidated(string? email)
        {
            return string.IsNullOrEmpty(email) || !IsEmailAddress(email);
        }

        private bool IsEmailAddress(string email)
        {
            string emailPattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            return Regex.IsMatch(email, emailPattern);
        }
    }
}

