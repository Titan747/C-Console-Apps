using Spire.Xls;
using System.Linq;

namespace RemoveDuplicateRows
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a Workbook instance
            Workbook workbook = new Workbook();
            //Load the Excel file
            workbook.LoadFromFile("C:\\Users\\addsi\\source\\Test.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Specify the range that you want to remove duplicate records from.
            var range = sheet.Range["A1:A" + sheet.LastRow];
            //Get the duplicated row numbers  
            var duplicatedRows = range.Rows
                   .GroupBy(x => x.Columns[0].DisplayedText)
                   .Where(x => x.Count() > 1)
                   .SelectMany(x => x.Skip(1))
                   .Select(x => x.Columns[0].Row)
                   .ToList();

            //Remove the duplicate rows & blank rows if any           
            for (int i = 0; i < duplicatedRows.Count; i++)
            {
                sheet.DeleteRow(duplicatedRows[i] - i);
            }

            //Save the result file
            Console.WriteLine("\n");
            Console.WriteLine("...Duplicates Removed Successfully...");
            workbook.SaveToFile("C:\\Users\\addsi\\source\\Output.xlsx", ExcelVersion.Version2013);
        }
    }
}