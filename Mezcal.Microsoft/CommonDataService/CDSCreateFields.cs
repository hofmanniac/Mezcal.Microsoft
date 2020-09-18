using Mezcal.Commands;
using Mezcal.Connections;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSCreateFields : ICommand
    {
        private CDSConnection _cdsConnection;

        public void Process(JObject command, Context context)
        {
            var config = command["config"];
            if (config == null) { return; }
            var joConfig = context.FindItemByName(config.ToString());

            this._cdsConnection = new CDSConnection(joConfig);
            if (this._cdsConnection == null) { return; }

            this.CreateFields(command, context);
        }

        private void CreateFields(JToken field, Context context)
        {
            var entityname = JSONUtil.GetText(field, "#cds-fields");
            if (entityname == null) { entityname = JSONUtil.GetText(field, "entityname"); }
            if (entityname == null) { return; }

            var joEntity = context.FindItemByName(entityname);
            if (joEntity == null) { return; }

            var fields = joEntity["items"];

            CDSCreateField cdsCreateField = new CDSCreateField();
            cdsCreateField.SetConnection(this._cdsConnection);

            foreach (var item in fields)
            {
                var joItem = (JObject)item;
                joItem.Add("entity", entityname);

                var disabled = JSONUtil.GetText(item, "disabled");
                if (disabled != null && disabled == "true") { continue; }

                cdsCreateField.CreateField(joItem);
            }
        }

        public JObject Prompt(CommandEngine commandEngine)
        {
            throw new NotImplementedException();
        }
    }
}
