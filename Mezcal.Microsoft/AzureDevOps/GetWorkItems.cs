using Mezcal.Commands;
using Microsoft.TeamFoundation.Common;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Mezcal.Microsoft.AzureDevOps
{
    public class GetWorkItems : ICommand
    {
        public void Process(JObject command, Context context)
        {
        }

        public void GetItems(string uri, string pat, string project, string setName, Context context)
        {
            var task = this.GetItems(uri, pat, project); task.Wait();
            var jaItems = this.ConvertToJSON(task.Result);
            //Console.WriteLine(workItems);

            context.Store(setName, jaItems);

            Console.WriteLine($"Retrieved work items from {uri}\\{project} into {setName}.");
        }

        private JArray ConvertToJSON(List<WorkItem> workItems)
        {
            var items = new JArray();

            foreach (var workitem in workItems)
            {
                var item = new JObject();
                item.Add("ID", workitem.Id);

                foreach (var field in workitem.Fields)
                {
                    //if (field.Key == "System.Title" || field.Key == "System.Parent" || field.Key == "System.WorkItemType" || field.Key == "Microsoft.VSTS.Scheduling.Effort" || field.Key == "System.Tags")
                    //{
                        item.Add(field.Key, field.Value.ToString());
                    //}
                }

                items.Add(item);
            }

            return items;
        }

        //private JObject ConvertToHierarchy(JArray items)
        //{
        //    var dict = new Dictionary<int, JToken>();

        //    foreach (var item in items)
        //    {
        //        var id = Int32.Parse(item["ID"].ToString());

        //        dict.Add(id, item);
        //    }

        //    foreach (var item in dict.Values)
        //    {
        //        if (item["System.Parent"] != null)
        //        {
        //            int parentNumber = Int32.Parse(item["System.Parent"].ToString());
        //            var parentItem = (JObject)dict[parentNumber];
        //            if (parentItem.ContainsKey("items") == false) { parentItem.Add("items", new JArray()); }
        //            var subItems = (JArray)parentItem["items"];
        //            subItems.Add(item);
        //        }
        //    }

        //    return (JObject)dict.FirstOrDefault().Value;
        //}

        private async Task<List<WorkItem>> GetItems(string uri, string pat, string project)
        {      
            var credentials = new VssBasicCredential(string.Empty, pat);

            // create a wiql object and build our query
            var wiql = new Wiql()
            {
                // NOTE: Even if other columns are specified, only the ID & URL will be available in the WorkItemReference
                Query = "Select [Id] " +
                        "From WorkItems " +
                        //"Where [Work Item Type] = 'Epic' " +
                        "Where [System.TeamProject] = '" + project + "'",
            };

            // create instance of work item tracking http client
            using (var httpClient = new WorkItemTrackingHttpClient(new Uri(uri), credentials))
            {
                // execute the query to get the list of work items in the results
                var result = await httpClient.QueryByWiqlAsync(wiql).ConfigureAwait(false);
                var ids = result.WorkItems.Select(item => item.Id).ToArray();

                // some error handling
                if (ids.Length == 0) { return null; }

                // build a list of the fields we want to see
                //var fields = new[] { "System.Id", "System.Title", "System.State, System.Parent" };
                //System.AreaPath
                //System.TeamProject
                //System.IterationPath
                //System.WorkItemType
                //System.State
                //System.Reason
                //System.CreatedDate
                //System.CreatedBy
                //System.ChangedDate
                //System.ChangedBy
                //System.CommentCount

                // get work items for the ids found in query
                //return await httpClient.GetWorkItemsAsync(ids, fields, result.AsOf).ConfigureAwait(false);
                return await httpClient.GetWorkItemsAsync(ids, expand: WorkItemExpand.Relations).ConfigureAwait(false);
            }
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }
    }
}
