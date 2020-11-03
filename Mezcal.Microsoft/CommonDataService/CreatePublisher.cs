using Mezcal;
using Mezcal.Commands;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerGov.Commands
{
    public class CreatePublisher : ICommand
    {
        public JObject Prompt(CommandEngine commandEngine)
        {
            return null;
        }

        public void Process(JObject command, Context context)
        {
            /*
            // Define a new publisher
            Publisher _myPublisher = new Publisher
            {
                UniqueName = "contoso-publisher",
                FriendlyName = "Contoso publisher",
                SupportingWebsiteUrl =
                  "https://docs.microsoft.com/powerapps/developer/common-data-service/overview",
                CustomizationPrefix = "contoso",
                EMailAddress = "someone@contoso.com",
                Description = "This publisher was created from sample code"
            };

            // Does the publisher already exist?
            QueryExpression querySamplePublisher = new QueryExpression
            {
                EntityName = Publisher.EntityLogicalName,
                ColumnSet = new ColumnSet("publisherid", "customizationprefix"),
                Criteria = new FilterExpression()
            };

            querySamplePublisher.Criteria.AddCondition("uniquename", ConditionOperator.Equal,
               _myPublisher.UniqueName);

            EntityCollection querySamplePublisherResults =
               _serviceProxy.RetrieveMultiple(querySamplePublisher);

            Publisher SamplePublisherResults = null;

            // If the publisher already exists, use it
            if (querySamplePublisherResults.Entities.Count > 0)
            {
                SamplePublisherResults = (Publisher)querySamplePublisherResults.Entities[0];
                _publisherId = (Guid)SamplePublisherResults.PublisherId;
                _customizationPrefix = SamplePublisherResults.CustomizationPrefix;
            }

            // If the publisher doesn't exist, create it
            if (SamplePublisherResults == null)
            {
                _publisherId = _serviceProxy.Create(_myPublisher);

                Console.WriteLine(String.Format("Created publisher: {0}.",
                _myPublisher.FriendlyName));

                _customizationPrefix = _myPublisher.CustomizationPrefix;
            }
            */
        }
    }
}
