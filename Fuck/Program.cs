namespace Fuck
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using ExcelDataReader;
    using System.IO;

    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\sfernandez\Downloads\List_Basement_All.xlsx";
            List<string> lines = new List<string>();
            List<string> emails = new List<string>();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            var oldEmail = reader.GetValue(1);
                            var newEmail = reader.GetValue(2);
                            var line = "WHEN email=\"" + oldEmail + "\" THEN \"" + newEmail + "\"";
                            if ((string)oldEmail != "-") {
                                lines.Add(line);
                            }
                            Console.WriteLine(line);
                            // emails.Add("('" + oldEmail + "'),");
                            // emails.Add("('" + newEmail + "'),");
                        }
                    } while (reader.NextResult());

                    var counter = 0;
                    File.WriteAllLines("file.sql", lines);
                    //File.WriteAllLines("email.sql", emails);
                    //StreamWriter file = new StreamWriter("file.sql");
                    //foreach (string line in lines)
                    //{
                    //    Console.WriteLine(counter++ + " == " + line);
                    //    // If the line doesn't contain the word 'Second', write the line to the file.
                    //    file.WriteLine(line);
                    //}
                    Console.ReadLine();
                }
            }
        }
    }
}
