using Mezcal.Commands;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using E = Microsoft.Office.Interop.Excel;

namespace Mezcal.Microsoft.Office
{
    public class LoadExcel : ICommand
    {
        public void Process(JObject command, Context context)
        {
            string file = command["file"].ToString();
            string set = command["into"].ToString();
            string worksheetName = command["worksheet"].ToString();
            int startColumn = Int32.Parse(command["startcol"].ToString());
            int endColumn = Int32.Parse(command["endcol"].ToString());
            int startRow = Int32.Parse(command["startrow"].ToString());
            int endRow = Int32.Parse(command["endrow"].ToString());

            Console.WriteLine("Loading Excel file {0} into {1}", file, set);
            var sub = this.ReadFlatTable(file, worksheetName, startColumn, endColumn, startRow, endRow);

            context.Store(set, sub);
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Assumes first row contains column headings
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="worksheetName"></param>
        /// <param name="startColumn"></param>
        /// <param name="endColumn"></param>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <returns></returns>
        public JArray ReadFlatTable(string filename, string worksheetName, int startColumn, int endColumn, int startRow, int endRow = -1)
        {
            JArray result = new JArray();

            var app = new E.Application();
            var workbook = app.Workbooks.Open(filename);
            E.Worksheet worksheet = null;

            foreach (E.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == worksheetName) { worksheet = sheet; break; }
            }

            if (worksheet == null) { return null; }

            Dictionary<int, string> columnNames = new Dictionary<int, string>();

            for (int column = startColumn; column <= endColumn; column++)
            {
                string cellValue = this.CellValue(worksheet, startRow, column);
                columnNames[column] = cellValue;
            }

            for (int row = startRow + 1; row < endRow; row++)
            {
                JObject rec = new JObject();

                for (int column = startColumn; column <= endColumn; column++)
                {
                    string cellValue = this.CellValue(worksheet, row, column);

                    string colName = columnNames[column];


                    rec.Add(colName, cellValue);
                    //rec.Fields.Add(colName, cellValue);

                }

                result.Add(rec);
            }

            workbook.Close();

            return result;
        }

        public string CellValue(E.Worksheet worksheet, int row, int column)
        {
            string result = null;

            if (column == 0) { return result; }

            var rawcell = worksheet.Cells[row, column];
            var val2 = ((E.Range)rawcell).Value2;
            if (val2 != null) { result = val2.ToString(); }

            return result;
        }
    }
}
