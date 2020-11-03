using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mezcal.Connections;
using System.Runtime.CompilerServices;

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSConnection : IConnection
    {
        private JObject _config = null;

        private IOrganizationService ServiceEndpoint = null;

        public CDSConnection(JToken config)
        {
            if (config.Type == JTokenType.Object)
            {
                this._config = (JObject)config;
            }
        }

        public static CDSConnection FromCommand(JToken command, Context context)
        {
            var config = command["env"];           
            var cdsConnection = new CDSConnection(config);
            if (cdsConnection == null) { Console.WriteLine("No environment configuration information available. Aborting operation."); return null; }
            return cdsConnection;
        }

        private IOrganizationService GetService()
        {
            StringBuilder connectionString = new StringBuilder();

            if (this._config["appid"] != null)
            {
                var url = _config["url"];
                var appid = _config["appid"];
                var secret = _config["clientsecret"];

                connectionString.Append($"AuthType=ClientSecret;");
                connectionString.Append($"url={url};");
                connectionString.Append($"ClientId={appid};");
                connectionString.Append($"ClientSecret={secret}");
            }
            else
            {
                var url = _config["url"];
                var password = _config["p"];
                var username = _config["username"];

                connectionString.Append($"AuthType=Office365;");
                connectionString.Append($"RequireNewInstance=True;");
                connectionString.Append($"Username = {username};");
                connectionString.Append($"Password = {password};");
                connectionString.Append($"Url = {url}");
            }

            Console.WriteLine($"Connecting to {_config["url"]}...");

            var result = GetService(connectionString.ToString());
            return result;
        }

        private IOrganizationService GetService(string connectionString)
        {
            var client = new CrmServiceClient(connectionString);

            if (client.OrganizationServiceProxy == null)
            {
                if (client.OrganizationWebProxyClient == null)
                {
                    Console.WriteLine("Unable to connect to Service: " + client.LastCrmError);
                    return null;
                }
                else
                {
                    return client.OrganizationWebProxyClient;
                }
            }
            else
            {
                var serviceProxy = client.OrganizationServiceProxy;
                serviceProxy.Timeout = new TimeSpan(0, 15, 0);
                //serviceProxy.Authenticate();
                //serviceProxy.EnableProxyTypes();
                IOrganizationService service = serviceProxy;
                return service;
            }
        }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            this.Connect();
            if (this.ServiceEndpoint == null) { return null; }

            return this.ServiceEndpoint.Execute(request);
        }

        public Entity Retrieve(string entityName, Guid id, ColumnSet columnSet)
        {
            this.Connect();

            try
            {
                return this.ServiceEndpoint.Retrieve(entityName, id, columnSet);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        public static List<Entity> ConvertToCDSEntities(string entityName, JArray data, JArray map, CDSConnection cdsConnection)
        {
            var result = new List<Entity>();

            var em = cdsConnection.GetEntityMetadata(entityName);
            if (em == null)
            {
                Console.WriteLine("Unable to load entity metadata for " + entityName + " - skipping.");
                return null;
            }

            foreach (var item in data)
            {
                Entity cdsEntity = new Entity(entityName);
                var record = (JObject)item;

                foreach (var field in record)
                {
                    string targetName = DetermineTargetName(field.Key, map, record);
                    if (targetName == null) { continue; }

                    var atr = em.Attributes.FirstOrDefault(a => a.LogicalName == targetName);
                    if (atr == null) { continue; }

                    var atrVal = ConvertToAttribute(atr, record[field.Key]);
                    if (atrVal == null) { continue; }
                    cdsEntity.Attributes.Add(targetName, atrVal);
                }

                result.Add(cdsEntity);
            }

            return result;
        }

        private static object ConvertToAttribute(AttributeMetadata atr, JToken valToken)
        {
            if (valToken == null) { return null; }

            if (atr.AttributeTypeName.Value == "LookupType")
            {
                string lookupTo = null;
                string id = null;

                if (valToken.HasValues == false)
                {
                    var erAtr = (LookupAttributeMetadata)atr;
                    lookupTo = erAtr.Targets[0];
                    id = valToken.ToString();
                }

                if (lookupTo != null && id != null)
                {
                    var er = new EntityReference(lookupTo);
                    er.Id = new Guid(id);
                    return er;
                }
            }
            else if (atr.AttributeTypeName.Value == "PicklistType")
            {
                string val = null;

                if (valToken.HasValues == false) { val = valToken.ToString(); }
                else { val = valToken["value"].ToString(); }

                int numval = Int32.Parse(val);
                var osv = new OptionSetValue(numval);
                return osv;
            }
            else if (atr.AttributeTypeName.Value == "IntegerType")
            {
                string val = valToken.ToString();
                int numval = Int32.Parse(val);
                return numval;
            }
            else if (atr.AttributeTypeName.Value == "DateTimeType")
            {
                string val = valToken.ToString();
                var valDateTime = DateTime.Parse(val);
                return valDateTime;
            }
            else if (atr.AttributeTypeName.Value == "BooleanType")
            {
                string val = valToken.ToString();
                var valBoolean = Boolean.Parse(val);
                return valBoolean;
            }
            else if (atr.AttributeTypeName.Value == "MoneyType")
            {
                string val = valToken.ToString();
                var valMoney = new Money(Decimal.Parse(val));
                return valMoney;
            }
            else if (atr.AttributeTypeName.Value == "UniqueidentifierType")
            {
                Guid val = new Guid(valToken.ToString());
                return val;
            }
            else
            {
                string val = valToken.ToString();
                return val;
            }

            return null;
        }

        private static string DetermineTargetName(string fieldName, JArray map, JObject record)
        {
            string targetName = fieldName;

            if (map != null)
            {
                // find this field in the map, lookup by source
                var mapEntry = map.FirstOrDefault(m => m["source"].ToString() == fieldName);

                // if no map entry found by source field, then not mapped - skip this field
                if (mapEntry == null) { return null; }

                // else a map entry was found for this source, grab target info
                var target = mapEntry["target"];

                // if target info not found - assume will use source name
                if (target == null) { targetName = mapEntry["source"].ToString(); }

                // target info was found
                else
                {
                    // if target is a simple string, use that
                    if (target.Type == JTokenType.String)
                    {
                        targetName = target.ToString();
                    }
                    else if (target.Type == JTokenType.Array)
                    {
                        // else is array of choices, need to figure out which one applies
                        var jaTarget = (JArray)target;

                        foreach (var targetOption in jaTarget)
                        {
                            // {"#if": {"equals": "govcdm_role", "value": "12345"}, "#then": "govcdm_primarysubject"}
                            var condition = targetOption["#if"];
                            var conditionField = condition["equals"];
                            var conditionValue = condition["value"];

                            var recordValue = record[conditionField.ToString()];
                            if (recordValue.ToString() == conditionValue.ToString())
                            {
                                var action = targetOption["#then"];
                                targetName = action.ToString();
                                break;
                            }

                        }
                    }
                }
            }

            return targetName;
        }
        public JArray RetrieveEntityData(string entity, ColumnSet columns)
        {
            this.Connect();

            JArray result = new JArray();

            var query = new QueryExpression(entity);
            query.ColumnSet = columns;

            var items = this.ServiceEndpoint.RetrieveMultiple(query);

            foreach (var item in items.Entities)
            {
                JObject record = new JObject();

                foreach (var atr in item.Attributes)
                {
                    string val = atr.Value.ToString();

                    if (atr.Value is EntityReference er)
                    {
                        val = er.Id.ToString();
                    }
                    else if (atr.Value is OptionSetValue osv)
                    {
                        val = osv.Value.ToString();
                    }

                    record.Add(atr.Key, val);
                }

                result.Add(record);
            }

            return result;
        }

        public JArray RetrieveOptionSet(string entityName, string fieldName)
        {
            this.Connect();

            var req = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = fieldName,
                RetrieveAsIfPublished = true

            };

            var attributeResponse = (RetrieveAttributeResponse)this.ServiceEndpoint.Execute(req);
            var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

            JArray result = new JArray();

            foreach (var option in attributeMetadata.OptionSet.Options)
            {
                JObject jOption = new JObject();
                jOption.Add("value", option.Value);
                jOption.Add("label", option.Label.UserLocalizedLabel.Label.ToString());
                result.Add(jOption);
            }

            return result;
        }

        public Guid Create(Entity entity)
        {
            this.Connect();
            return this.ServiceEndpoint.Create(entity);
        }

        public void Update(Entity entity)
        {
            this.Connect();
            this.ServiceEndpoint.Update(entity);

        }
        private void Connect()
        {
            if (this.ServiceEndpoint == null) { this.ServiceEndpoint = GetService(); }
        }

        public EntityMetadata GetEntityMetadata(string entityName)
        {
            RetrieveEntityRequest request = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = entityName
            };

            try
            {
                var response = (RetrieveEntityResponse)this.Execute(request);
                if (response == null) { return null; }

                EntityMetadata entityMetaData = response.EntityMetadata;
                return entityMetaData;
            }
            catch
            {
                return null;
            }
        }
    }
}
