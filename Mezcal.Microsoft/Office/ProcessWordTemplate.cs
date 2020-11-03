using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Mezcal.Commands;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mezcal.Microsoft.Office
{
    public class ProcessWordTemplate : ICommand
    {
        private MemoryStream _templateStream = null;
        private Context _context;
        private JArray _fieldMap = null;

        public void Process(JObject command, Context context)
        {

            throw new NotImplementedException();
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }

        private void PopulateDocument()
        {           
            _context.Trace("Populate Document");

            // process the document
            var doc = WordprocessingDocument.Open(this._templateStream, true);
            var body = doc.MainDocumentPart.Document.Body;

            // process all the document elements
            this.ProcessElements(body);

            // commit, close out the doc
            doc.Close();
        }

        /// <summary>
        /// Recursive function to process all elements.
        /// Looks for Structured Document Tags to process.
        /// </summary>
        /// <param name="element"></param>
        private void ProcessElements(OpenXmlElement element)
        {
            foreach (var subElement in element.ChildElements)
            {
                if (subElement.LocalName == "sdt")
                {
                    _context.Trace($"Processing Structured Document Tag {subElement.LocalName}");

                    this.ProcessStdBlock(subElement);
                }

                this.ProcessElements(subElement);
            }
        }

        private void ProcessStdBlock(OpenXmlElement element)
        {
            JObject field = null;

            foreach (var subElement in element.ChildElements)
            {
                if (subElement.LocalName == "sdtPr")
                {
                    field = this.ProcessStdProperties(subElement);
                }
                else if (subElement.LocalName == "sdtContent")
                {
                    // figure out what field we should look for
                    var fieldName = field["alias"].ToString();

                    // look for that field in the data
                    //var jtValue = this._data[fieldName];
                    var jtValue = this.ResolveAliasToValue(fieldName);

                    if (jtValue != null)
                    {
                        Text placeHolderText = this.LocatePlaceholder(subElement);
                        if (placeHolderText == null) { _context.Trace("Warning - cannot find placeholder text"); continue; }
                        _context.Trace($"Replacing {fieldName} with {jtValue}");
                        placeHolderText.Text = jtValue.ToString();
                    }
                }
            }
        }

        private Text LocatePlaceholder(OpenXmlElement element)
        {
            Text result = null;

            foreach (var contentElement in element.ChildElements)
            {
                if (contentElement.LocalName == "p")
                {
                    result = this.LocatePlaceholder(contentElement);
                }
                else if (contentElement.LocalName == "r")
                {
                    result = this.LocatePlaceholder(contentElement);
                }
                else if (contentElement.LocalName == "t")
                {
                    result = (Text)contentElement;
                }

                if (result != null) { break; }
            }

            return result;
        }

        private JObject ProcessStdProperties(OpenXmlElement element)
        {
            var result = new JObject();

            foreach (var prop in element.ChildElements)
            {
                if (prop.LocalName == "alias")
                {
                    SdtAlias sdtAlias = (SdtAlias)prop;
                    result.Add("alias", sdtAlias.Val.ToString());
                }
                else if (prop.LocalName == "tag")
                {
                    Tag t = (Tag)prop;
                    result.Add("tag", t.Val.ToString());
                }
            }

            return result;
        }

        private JToken ResolveAliasToValue(string alias)
        {
            foreach (JObject joMapItem in this._fieldMap)
            {
                if (joMapItem["field"].ToString() == alias)
                {
                    var source = joMapItem["source"];
                    var newValue = this.Resolve(source.ToString());
                    return newValue;
                }
            }

            return null;
        }

        private JToken Resolve(JToken command, string arg)
        {
            var argValue = command[arg];
            if (argValue == null) { return null; }

            return this.Resolve(argValue.ToString());
        }

        private JToken Resolve(string value)
        {
            var result = value;

            // mezcal - this needs looked at
            if (result.Equals("@id")) { result = _context.Variables["id"].ToString(); }

            if (result.StartsWith("$"))
            {
                var parts = result.Split('.');
                var set = parts[0].Substring(1);
                var field = parts[1];
                var sourceSet = (JToken)_context.Fetch(set);
                var sourceValue = sourceSet[field];
                if (sourceValue != null) { result = sourceValue.ToString(); }
            }

            return result;
        }
    }
}
