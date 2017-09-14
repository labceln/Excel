using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ClosedXML.Excel;
using System.IO;
using LumenWorks.Framework.IO.Csv;

namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
       

        private void button1_Click(object sender, EventArgs e)
        {


            Data_Type();
            Showcase();
            /*
            var x = new double[] { 3, 4, 5, 6, 7 };
            var y = new double[] { 2, 1, 3, 5, 6 };
            var z = new double[] { 3, 1, 0, -3, 4 };
            try
            {
                // appendをtrueにすると，既存のファイルに追記
                //         falseにすると，ファイルを新規作成する
                var append = false;
                // 出力用のファイルを開く
                using (var sw = new System.IO.StreamWriter(@"test.csv", append))
                {
                    for (int i = 0; i < x.Length; ++i)
                    {
                        // 
                        sw.WriteLine("{0}, {1}, {2},", x[i], y[i], z[i]);
                    }
                }
            }
            catch (System.Exception ex)
            {
                // ファイルを開くのに失敗したときエラーメッセージを表示
                System.Console.WriteLine(ex.Message);
            }
            */
        }
        private void Showcase()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Contacts");

            // Title
            ws.Cell("B2").Value = "Contacts";

            // First Names
            ws.Cell("B3").Value = "FName";
            ws.Cell("B4").Value = "John";
            ws.Cell("B5").Value = "Hank";
            ws.Cell("B6").SetValue("Dagny"); // Another way to set the value

            // Last Names
            ws.Cell("C3").Value = "LName";
            ws.Cell("C4").Value = "Galt";
            ws.Cell("C5").Value = "Rearden";
            ws.Cell("C6").SetValue("Taggart"); // Another way to set the value

            // Boolean
            ws.Cell("D3").Value = "Outcast";
            ws.Cell("D4").Value = true;
            ws.Cell("D5").Value = false;
            ws.Cell("D6").SetValue(false); // Another way to set the value

            // DateTime
            ws.Cell("E3").Value = "DOB";
            ws.Cell("E4").Value = new DateTime(1919, 1, 21);
            ws.Cell("E5").Value = new DateTime(1907, 3, 4);
            ws.Cell("E6").SetValue(new DateTime(1921, 12, 15)); // Another way to set the value

            // Numeric
            ws.Cell("F3").Value = "Income";
            ws.Cell("F4").Value = 2000;
            ws.Cell("F5").Value = 40000;
            ws.Cell("F6").SetValue(10000); // Another way to set the value


            // From worksheet
            var rngTable = ws.Range("B2:F6");

            // From another range
            var rngDates = rngTable.Range("D3:D5"); // The address is relative to rngTable (NOT the worksheet)
            var rngNumbers = rngTable.Range("E3:E5"); // The address is relative to rngTable (NOT the worksheet)

            // Using a OpenXML's predefined formats
            rngDates.Style.NumberFormat.NumberFormatId = 15;

            // Using a custom format
            rngNumbers.Style.NumberFormat.Format = "$ #,##0";

            rngTable.FirstCell().Style
                .Font.SetBold()
                .Fill.SetBackgroundColor(XLColor.CornflowerBlue)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

            rngTable.FirstRow().Merge(); // We could've also used: rngTable.Range("A1:E1").Merge() or rngTable.Row(1).Merge()

            var rngHeaders = rngTable.Range("A2:E2"); // The address is relative to rngTable (NOT the worksheet)
            rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngHeaders.Style.Font.Bold = true;
            rngHeaders.Style.Font.FontColor = XLColor.DarkBlue;
            rngHeaders.Style.Fill.BackgroundColor = XLColor.Aqua;

            var rngData = ws.Range("B3:F6");
            var excelTable = rngData.CreateTable();

            // Add the totals row
            excelTable.ShowTotalsRow = true;
            // Put the average on the field "Income"
            // Notice how we're calling the cell by the column name
            excelTable.Field("Income").TotalsRowFunction = XLTotalsRowFunction.Average;
            // Put a label on the totals cell of the field "DOB"
            excelTable.Field("DOB").TotalsRowLabel = "Average:";

            // Add thick borders to the contents of our spreadsheet
            ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            // You can also specify the border for each side:
            // contents.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            // contents.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thick;
            // contents.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thick;
            // contents.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

            ws.Columns().AdjustToContents(); // You can also specify the range of columns to adjust, e.g.
                                             // ws.Columns(2, 6).AdjustToContents(); or ws.Columns("2-6").AdjustToContents();

            wb.SaveAs("Showcase.xlsx");


        }
        private void Csv()
        {
            using (CsvReader csv =
                new CsvReader(new StreamReader("test.csv"), false))
            {
                int fieldCount = csv.FieldCount;

                //string[] headers = csv.GetFieldHeaders();
                while (csv.ReadNextRecord())
                {
                    for (int i = 0; i < fieldCount; i++)
                        Console.Write(string.Format("{0};",
                                       csv[i]));
                    Console.WriteLine();
                }
            }
        }

        private void Data_Type()
        {

            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Data Types");

            var co = 2;
            var ro = 1;

            ws.Cell(++ro, co).Value = "Plain Text:";
            ws.Cell(ro, co + 1).Value = "Hello World.";

            ws.Cell(++ro, co).Value = "Plain Date:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);

            ws.Cell(++ro, co).Value = "Plain DateTime:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2, 13, 45, 22);

            ws.Cell(++ro, co).Value = "Plain Boolean:";
            ws.Cell(ro, co + 1).Value = true;

            ws.Cell(++ro, co).Value = "Plain Number:";
            ws.Cell(ro, co + 1).Value = 123.45;

            ws.Cell(++ro, co).Value = "TimeSpan:";
            ws.Cell(ro, co + 1).Value = new TimeSpan(33, 45, 22);

            ro++;

            ws.Cell(++ro, co).Value = "Explicit Text:";
            ws.Cell(ro, co + 1).Value = "'Hello World.";

            ws.Cell(++ro, co).Value = "Date as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2).ToString();

            ws.Cell(++ro, co).Value = "DateTime as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2, 13, 45, 22).ToString();

            ws.Cell(++ro, co).Value = "Boolean as Text:";
            ws.Cell(ro, co + 1).Value = "'" + true.ToString();

            ws.Cell(++ro, co).Value = "Number as Text:";
            ws.Cell(ro, co + 1).Value = "'123.45";

            ws.Cell(++ro, co).Value = "TimeSpan as Text:";
            ws.Cell(ro, co + 1).Value = "'" + new TimeSpan(33, 45, 22).ToString();

            ro++;

            ws.Cell(++ro, co).Value = "Changing Data Types:";

            ro++;

            ws.Cell(++ro, co).Value = "Date to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "DateTime to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2, 13, 45, 22);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Boolean to Text:";
            ws.Cell(ro, co + 1).Value = true;
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Number to Text:";
            ws.Cell(ro, co + 1).Value = 123.45;
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "TimeSpan to Text:";
            ws.Cell(ro, co + 1).Value = new TimeSpan(33, 45, 22);
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Text to Date:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.DateTime;

            ws.Cell(++ro, co).Value = "Text to DateTime:";
            ws.Cell(ro, co + 1).Value = "'" + new DateTime(2010, 9, 2, 13, 45, 22).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.DateTime;

            ws.Cell(++ro, co).Value = "Text to Boolean:";
            ws.Cell(ro, co + 1).Value = "'" + true.ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.Boolean;

            ws.Cell(++ro, co).Value = "Text to Number:";
            ws.Cell(ro, co + 1).Value = "'123.45";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Number;

            ws.Cell(++ro, co).Value = "Text to TimeSpan:";
            ws.Cell(ro, co + 1).Value = "'" + new TimeSpan(33, 45, 22).ToString();
            ws.Cell(ro, co + 1).DataType = XLCellValues.TimeSpan;

            ro++;

            ws.Cell(++ro, co).Value = "Formatted Date to Text:";
            ws.Cell(ro, co + 1).Value = new DateTime(2010, 9, 2);
            ws.Cell(ro, co + 1).Style.DateFormat.Format = "yyyy-MM-dd";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ws.Cell(++ro, co).Value = "Formatted Number to Text:";
            ws.Cell(ro, co + 1).Value = 12345.6789;
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;

            ro++;

            ws.Cell(++ro, co).Value = "Blank Text:";
            ws.Cell(ro, co + 1).Value = 12345.6789;
            ws.Cell(ro, co + 1).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(ro, co + 1).DataType = XLCellValues.Text;
            ws.Cell(ro, co + 1).Value = "";

            ro++;

            // Using inline strings (few users will ever need to use this feature)
            //
            // By default all strings are stored as shared so one block of text
            // can be reference by multiple cells.
            // You can override this by setting the .ShareString property to false
            ws.Cell(++ro, co).Value = "Inline String:";
            var cell = ws.Cell(ro, co + 1);
            cell.Value = "Not Shared";
            cell.ShareString = false;

            // To view all shared strings (all texts in the workbook actually), use the following:
            // workbook.GetSharedStrings()

            ws.Columns(2, 3).AdjustToContents();

            workbook.SaveAs("DataTypes.xlsx");

        }
    }
}
