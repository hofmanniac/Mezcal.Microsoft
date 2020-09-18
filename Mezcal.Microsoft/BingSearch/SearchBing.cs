using Mezcal.Commands;
using Mezcal.Connections;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Mezcal.Microsoft.BingSearch
{
    public class SearchBing : ICommand
    {
        private string accessKey = null;

        public void Process(JObject command, Context context)
        {
            context.Trace("bing-websearch");

            var uriKey = "$bingsearch-uri";
            if (context.Variables.ContainsKey(uriKey) == false) { return; }
            var uriBase = context.Variables[uriKey].ToString();

            var configKey = "$bingsearch-key";
            if (context.Variables.ContainsKey(configKey) == false) { return; }
            accessKey = context.Variables[configKey].ToString();

            var searchTerm = JSONUtil.GetText(command, "#bing-websearch");
            if (searchTerm == null) { searchTerm = JSONUtil.GetText(command, "search-term"); }

            Console.OutputEncoding = Encoding.UTF8;
            if (accessKey.Length == 32)
            {
                this.ConsoleWriteLine("Searching the Web for: " + searchTerm);
                SearchResult result = BingWebSearch(uriBase, searchTerm);
                //this.ConsoleWriteLine("\nRelevant HTTP Headers:\n");
                //foreach (var header in result.relevantHeaders)
                //    Console.WriteLine(header.Key + ": " + header.Value);

                //this.ConsoleWriteLine("\nJSON Response:\n");
                //Console.WriteLine(JsonPrettyPrint(result.jsonResult));
                JToken jResult = JToken.Parse(result.jsonResult);
                //this.ConsoleWriteLine(jResult.ToString());

                var setName = JSONUtil.GetText(command, "set");
                if (setName != null) { context.Store(setName, jResult); }

            }
            else
            {
                this.ConsoleWriteLine("Invalid Bing Search API subscription key!");
                this.ConsoleWriteLine("Please paste yours into the source code.");
            }

        }

        private void ConsoleWriteLine(string text)
        {
            Console.WriteLine(text);
        }

        struct SearchResult
        {
            public String jsonResult;
            public Dictionary<String, String> relevantHeaders;
        }

        private SearchResult BingWebSearch(string uriBase, string searchQuery)
        {
            // Construct the search request URI.
            var uriQuery = uriBase + "?q=" + Uri.EscapeDataString(searchQuery);

            // Perform request and get a response.
            WebRequest request = HttpWebRequest.Create(uriQuery);
            request.Headers["Ocp-Apim-Subscription-Key"] = accessKey;
            HttpWebResponse response = (HttpWebResponse)request.GetResponseAsync().Result;
            string json = new StreamReader(response.GetResponseStream()).ReadToEnd();

            // Create a result object.
            var searchResult = new SearchResult()
            {
                jsonResult = json,
                relevantHeaders = new Dictionary<String, String>()
            };

            // Extract Bing HTTP headers.
            foreach (String header in response.Headers)
            {
                if (header.StartsWith("BingAPIs-") || header.StartsWith("X-MSEdge-"))
                    searchResult.relevantHeaders[header] = response.Headers[header];
            }
            return searchResult;
        }

        ///// <summary>
        ///// Formats the JSON string by adding line breaks and indents.
        ///// </summary>
        ///// <param name="json">The raw JSON string.</param>
        ///// <returns>The formatted JSON string.</returns>
        //static string JsonPrettyPrint(string json)
        //{
        //    if (string.IsNullOrEmpty(json))
        //        return string.Empty;

        //    json = json.Replace(Environment.NewLine, "").Replace("\t", "");

        //    StringBuilder sb = new StringBuilder();
        //    bool quote = false;
        //    bool ignore = false;
        //    char last = ' ';
        //    int offset = 0;
        //    int indentLength = 2;

        //    foreach (char ch in json)
        //    {
        //        switch (ch)
        //        {
        //            case '"':
        //                if (!ignore) quote = !quote;
        //                break;
        //            case '\\':
        //                if (quote && last != '\\') ignore = true;
        //                break;
        //        }

        //        if (quote)
        //        {
        //            sb.Append(ch);
        //            if (last == '\\' && ignore) ignore = false;
        //        }
        //        else
        //        {
        //            switch (ch)
        //            {
        //                case '{':
        //                case '[':
        //                    sb.Append(ch);
        //                    sb.Append(Environment.NewLine);
        //                    sb.Append(new string(' ', ++offset * indentLength));
        //                    break;
        //                case ']':
        //                case '}':
        //                    sb.Append(Environment.NewLine);
        //                    sb.Append(new string(' ', --offset * indentLength));
        //                    sb.Append(ch);
        //                    break;
        //                case ',':
        //                    sb.Append(ch);
        //                    sb.Append(Environment.NewLine);
        //                    sb.Append(new string(' ', offset * indentLength));
        //                    break;
        //                case ':':
        //                    sb.Append(ch);
        //                    sb.Append(' ');
        //                    break;
        //                default:
        //                    if (quote || ch != ' ') sb.Append(ch);
        //                    break;
        //            }
        //        }
        //        last = ch;
        //    }
        //    return sb.ToString().Trim();
        //}

        public JObject Prompt(CommandEngine commandEngine)
        {
            JObject jCommand = new JObject();
            jCommand.Add("command", "bing-websearch");

            commandEngine.PromptForArgument(jCommand, "search-term");
            commandEngine.PromptForArgument(jCommand, "set");

            return jCommand;
        }
    }
}
