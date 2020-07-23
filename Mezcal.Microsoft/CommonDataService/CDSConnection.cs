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

namespace Mezcal.Microsoft.CommonDataService
{
    public class CDSConnection : IConnection
    {
        private ConnectionConfig _config = null;

        private IOrganizationService ServiceEndpoint = null;

        public CDSConnection(ConnectionConfig config)
        {
            this._config = config;
        }

        private IOrganizationService GetService()
        {
            StringBuilder connectionString = new StringBuilder();
            connectionString.Append("AuthType=Office365; RequireNewInstance=True;");
            connectionString.Append("Username = " + _config.UserName + ";");
            connectionString.Append("Password = " + _config.Password + ";");
            connectionString.Append("Url = " + _config.Url);

            Console.WriteLine("Connecting to " + _config.Url + "...");
            var result = GetService(connectionString.ToString());
            return result;
        }

        private IOrganizationService GetService(string connectionString)
        {
            CrmServiceClient client = new CrmServiceClient(connectionString);
            OrganizationServiceProxy serviceProxy = client.OrganizationServiceProxy;
            serviceProxy.Timeout = new TimeSpan(0, 15, 0);
            //serviceProxy.Authenticate();
            //serviceProxy.EnableProxyTypes();
            IOrganizationService service = serviceProxy;

            return service;
        }

        public OrganizationResponse Execute(OrganizationRequest request)
        {
            this.Connect();

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
