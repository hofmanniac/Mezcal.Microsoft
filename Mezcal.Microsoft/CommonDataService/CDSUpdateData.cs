using Mezcal.Commands;
using Mezcal.Connections;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSUpdateData : ICommand
    {
        public void Process(JObject command, Context context)
        {
            var entity = command["#cds-update-data"] ?? command["entity"];
            var source = command["source"];
            var map = command["map"];
            var idfield = command["id-field"];

            if (entity == null) return;
            if (source == null) return;
            if (idfield == null) return;

            CDSConnection cdsConnection = CDSConnection.FromCommand(command, context);
            if (cdsConnection == null) return;

            JArray set = (JArray)context.Fetch(source.ToString());

            var entities = CDSConnection.ConvertToCDSEntities(entity.ToString(), set, (JArray)map, cdsConnection);
            this.UpdateRecords(entities, idfield.ToString(), cdsConnection);
        }

        public void UpdateRecords(List<Entity> cdsEntities, string idfield, CDSConnection cdsConnection)
        {
            foreach (var item in cdsEntities)
            {
                //Entity cdsEntity = new Entity(item.LogicalName);

                if (item.Contains(idfield) == false) { continue; }

                var id = item[idfield].ToString();
                item.Id = Guid.Parse(id);

                Console.Write("Updating record...");
                cdsConnection.Update(item);
                Console.WriteLine("Updated.");
            }
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }
    }
}
