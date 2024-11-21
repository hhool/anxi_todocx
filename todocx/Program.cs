using System;

/// @note add description of app
/// read data from csv to work template
namespace todocx
{
    class Program
    {
        /// > parse args 
        /// e.g. todocx.exe -j input.json -i input.csv -t template.docx -o output.docx
        /// @param args command line arguments
        /// @return 0 if success, otherwise error code.
        static void Main(string[] args)
        {
            String[] arguments = Environment.GetCommandLineArgs();
            if (arguments.Length < 9)
            {
                Console.WriteLine("Usage: todocx.exe -j input.json -i input.csv -t template.docx -o output.docx");
                return;
            }
            string json = "";
            string csv = "";
            string template = "";
            string output = "";
            for (int i = 1; i < arguments.Length; i++)
            {
                if (arguments[i] == "-j")
                {
                    json = arguments[i + 1];
                }
                else if (arguments[i] == "-i")
                {
                    csv = arguments[i + 1];
                }
                else if (arguments[i] == "-t")
                {
                    template = arguments[i + 1];
                }
                else if (arguments[i] == "-o")
                {
                    output = arguments[i + 1];
                }
            }
            Console.WriteLine("json: " + json);
            Console.WriteLine("csv: " + csv);
            Console.WriteLine("template: " + template);
            Console.WriteLine("output: " + output);
            /// hide console window, run in background

            /// use docx template to generate docx
            Csv2Docx csv2Docx = new Csv2Docx();
            csv2Docx.GenerateDocx(json, csv, template, output);
            Console.WriteLine("Done!");
            // Go to http://aka.ms/dotnet-get-started-console to continue learning how to build a console app! 
        }
    }
}


