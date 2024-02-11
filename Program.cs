namespace GenerateEmlFile
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var mail = new EmlFileGenerator();
            mail.CreateEmlFile(args);
        }
    }
}