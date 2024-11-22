using System;
using System.Runtime.InteropServices;

/// @note add description of app
/// read data from csv to work template
namespace todocx
{

    class Program
    {

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
 
        /// Go to http://aka.ms/dotnet-get-started-console to continue learning how to build a console app!
        /// > parse args 
        /// e.g. todocx.exe -s summary.xml -i datalist.csv -t template.docx
        /// @param args command line arguments
        /// @return 0 if success, otherwise error code.
        static int Main(string[] args)
        {

            /// default hide todocx console window run in background.
            bool isConsoleVisible = false; // Flag to track console visibility
                                           // Toggle console visibility based on flag
            IntPtr hWndConsole = GetConsoleWindow();
            if (hWndConsole != IntPtr.Zero)
            {
                ShowWindow(hWndConsole, isConsoleVisible ? SW_SHOW : SW_HIDE);
            }
            String[] arguments = Environment.GetCommandLineArgs();
            if (arguments.Length < 7)
            {
                Console.WriteLine("Usage: todocx.exe -s summary.xml -i input.csv -t template.docx -o output.docx");
                return -1;
            }
            string xml = "";
            string csv = "";
            string template = "";
            for (int i = 1; i < arguments.Length; i++)
            {
                if (arguments[i] == "-s")
                {
                    xml = arguments[i + 1];
                }
                else if (arguments[i] == "-i")
                {
                    csv = arguments[i + 1];
                }
                else if (arguments[i] == "-t")
                {
                    template = arguments[i + 1];
                }
                else if (arguments[i] == "-f")
                {
                    isConsoleVisible = true;
                }
            }
            if (hWndConsole != IntPtr.Zero)
            {
                ShowWindow(hWndConsole, isConsoleVisible ? SW_SHOW : SW_HIDE);
            }
            Console.WriteLine("xml: " + xml + " csv: " + csv + " template: " + template);
            /// use docx template to generate docx
            Csv2Docx csv2Docx = new Csv2Docx();
            int ret = csv2Docx.GenerateDocx(xml, csv, template);
            return ret;
        }
    }
}


