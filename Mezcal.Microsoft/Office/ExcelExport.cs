using Mezcal.Commands;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using E = Microsoft.Office.Interop.Excel;

namespace Mezcal.Microsoft.Office
{
    public class ExcelExport : ICommand
    {
        public void Process(JObject command, Context context)
        {
            var setname = command["#excel-export"].ToString();
            var file = command["file"].ToString();

            this.Export(setname, file, context);
        }

        public void Export(string setname, string file, Context context)
        {
            var app = new E.Application();
            E.Workbook workbook = null;
            E.Worksheet worksheet = null;
            app.Visible = false;
            workbook = app.Workbooks.Add(E.XlWBATemplate.xlWBATWorksheet);

            try
            {
                worksheet = workbook.Worksheets[1]; // Compulsory Line in which sheet you want to write data

                var set = (JArray)context.Fetch(setname.ToString());

                int row = 1;

                foreach (JObject item in set)
                {
                    int col = 1;

                    foreach (var prop in item)
                    {
                        worksheet.Cells[row, col] = prop.Value.ToString();
                        col++;
                    }

                    row++;
                }

                //Writing data into excel of 100 rows with 10 column 
                //for (int r = 1; r < 101; r++) //r stands for ExcelRow and c for ExcelColumn
                //{
                //    // Excel row and column start positions for writing Row=1 and Col=1
                //    for (int c = 1; c < 11; c++)
                //        worksheet.Cells[r, c] = "R" + r + "C" + c;
                //}

                //workbook.Worksheets[1].Name = "MySheet"; //Renaming the Sheet1 to MySheet

                workbook.SaveAs(file.ToString());
                workbook.Close();
                app.Quit();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(app);
            }
            catch (Exception exHandle)
            {
                Console.WriteLine("Exception: " + exHandle.Message);
                Console.ReadLine();
            }
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }
    }
}
