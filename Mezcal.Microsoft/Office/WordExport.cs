using Mezcal.Commands;
using Mezcal.Connections;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace Mezcal.Microsoft.Office
{
    public class WordExport : ICommand
    {
        public void Process(JObject command, Context context)
        {
            // https://www.c-sharpcorner.com/UploadFile/muralidharan.d/how-to-create-word-document-using-C-Sharp/#:~:text=How%20to%20Create%20Word%20Document%20Using%20C%23.%201,executed%20successfully%2C%20the%20document%20output%20will%20be%3A%20

            var file = command["#word-export"];
            if (file == null) { file = command["file"]; }
            if (file == null) { return; }
            file = context.ReplaceVariables(file);

            context.Trace("Starting Microsoft Word...");
            Word.Application winword = new Word.Application(); 
            winword.ShowAnimation = false; 
            winword.Visible = false;
            object missing = System.Reflection.Missing.Value;

            //Create a new document  
            context.Trace("Creating new document...");
            Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            document.Content.SetRange(0, 0);

            context.Trace("Adding items...");
            var items = command["content"];

            foreach(var item in items)
            {
                context.Trace($"Adding item {item}");

                this.TryBlock(item, context, document);
            }

            object filename = file.ToString();
            context.Trace($"Items added. Saving document as...{file}");
            document.SaveAs2(ref filename);

            context.Trace("Opening document...");
            winword.Visible = true;

            //context.Trace("Document saved. Closing Word...");
            //document.Close(ref missing, ref missing, ref missing);
            //document = null;
            //winword.Quit(ref missing, ref missing, ref missing);
            //winword = null;
            //context.Trace("Word closed. Operation complete.");
        }

        private void TryBlock(JToken item, Context context, Word.Document document)
        {
            var block = item["block"];
            if (block == null) { return; }

            block = context.ReplaceVariables(block);

            object missing = System.Reflection.Missing.Value;
            Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            para1.Range.Text = block.ToString();

            if (item["style"] != null)
            {
                object styleHeading1 = item["style"].ToString();
                para1.Range.set_Style(ref styleHeading1);
            }

            var alignment = item["alignment"];
            if (alignment != null)
            {
                var sAlignment = alignment.ToString();
                if (sAlignment == "center")
                {
                    para1.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }

            para1.Range.InsertParagraphAfter();
        }

        private void TryLineBreak(JToken item, Context context, Word.Document document)
        {
            var linebreak = item["linebreak"];
            if (linebreak == null) { return; }

            object missing = System.Reflection.Missing.Value;
            Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
            para1.Range.Text = "\r\n";

            para1.Range.InsertParagraphAfter();

        }
        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }
    }
}
