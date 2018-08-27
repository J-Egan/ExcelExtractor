using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExtractor
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            /* 
             * Jack Egan
             * jackegan248@gmail.com
             * Excel data extractor
             */

            Console.Title = "Excel Data Extractor";

            Console.ForegroundColor = ConsoleColor.White;

            //Warn user each time the program starts
            Console.WriteLine("Please make sure that the excel file you want to search is");
            Console.WriteLine("1) In the same directory/folder as this program");
            Console.WriteLine("2) Called \"excelsheet.xlsx\"");
            Console.WriteLine();
            Console.WriteLine("Please make sure that the file \"output_Please_Rename.csv\" is deleted/renamed");
            Console.WriteLine();
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
            Console.Clear();

            //Load Excel file and start excel background process
            Console.WriteLine("Loading Excel File");
            string xlFile = Directory.GetCurrentDirectory() + "\\excelsheet.xlsx";

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(xlFile);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //decide if the user wants to search entire file or just n times
            //gather search conditions from user
            bool searchAll = false;
            int searchRange = 0;
            string searchText = "";

            Console.Clear();
            DisplayMenu(ref searchAll,ref searchRange,ref searchText);

            //decide what is going to be searched
            if (searchAll)
            {
                searchRange = xlWorkbook.Sheets.Count;
            }
            else
            {
                if (searchRange > xlWorkbook.Sheets.Count)
                {
                    Console.WriteLine("Correcting the entered amount of sheets");
                    Console.WriteLine("as there are only " + xlWorkbook.Sheets.Count + " sheets available");
                    searchRange = xlWorkbook.Sheets.Count;
                }
            }

            //initialise varibles for sheet data
            List<string> matchRows = new List<string>();
            string cell = "";
            string row = "";

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            double percent = 0;

            for (int l = 1; l <= searchRange; l++)
            {
                //update the current data set
                xlWorksheet = xlWorkbook.Sheets[l];
                xlRange = xlWorksheet.UsedRange;

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            //Take cell data and store it as a CSV style string
                            cell = Convert.ToString(xlRange.Cells[i, j].Value2.ToString() + "\t");
                            row = row + "," + cell;
                        }
                    }
                    //Check if the current row contains the text the user wants to search for
                    if (row.Contains(searchText))
                    {
                        //if the text is found, add the sheet title to the front and add it to the list of text files
                        row = xlWorksheet.Name + row;
                        matchRows.Add(row);
                    }

                    //reset the row for the next iteration
                    row = "";
                }

                //Update the sheet count and percentage each time
                Console.Clear();
                Console.WriteLine(l + " / " + searchRange + " sheets completed");

                percent = l;
                percent = percent / searchRange;
                percent *= 100;
                percent = Math.Round(percent, 2);

                Console.WriteLine(percent + "% complete");
            }

            //Open a textwriter stream
            TextWriter tw = new StreamWriter("output_Please_Rename.csv");

            //dump all matches to the file "output.csv"
            foreach (String s in matchRows)
            {
                tw.WriteLine(s);
            }

            //close text stream
            tw.Close();

            //prompt for excel issues closing
            Console.WriteLine();
            Console.WriteLine("Please look for an excel prompt and click on Don\'t Save");
            Console.WriteLine("You may need to minimise this window");

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            foreach (Excel.Worksheet sheet in xlWorkbook.Sheets)
            {
                Marshal.ReleaseComObject(sheet);
            }

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        public static void DisplayMenu(ref bool searchAll, ref int searchRange, ref string searchText) {

            Console.WriteLine("Do you want to search the entire file?");
            Console.WriteLine("Enter Y/N");
            string userInput = Convert.ToString(Console.ReadLine()).ToLower();
            Console.Clear();

            if (userInput == "n")
            {
                Console.WriteLine("Enter the amount of sheets you want to search through");
                searchRange = Convert.ToInt16(Console.ReadLine());
                searchAll = false;
            }
            else
            {
                searchAll = true;
            }

            Console.Clear();
            Console.WriteLine("Enter the text you want to search for");
            Console.WriteLine("This text needs to be exact, case sensitive and must include any text modifiers");
            searchText = Console.ReadLine();
            Console.Clear();
        }
    }
}