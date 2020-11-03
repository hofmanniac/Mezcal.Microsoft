using Mezcal;
using Mezcal.Commands;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Organization;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerGov.Commands
{
    public class CreateSolution : ICommand
    {
        public JObject Prompt(CommandEngine commandEngine)
        {
            return null;
        }

        public void Process(JObject command, Context context)
        {
            /*
            // Create a solution
            Solution solution = new Solution
            {
                SolutionUniqueName = "sample-solution",
                FriendlyName = "Sample solution",
                PublisherId = new EntityReference(Publisher.EntityLogicalName, _publisherId),
                //Description = "This solution was created by sample code.",
                VersionNumber = "1.0"
            };

            
            //Check whether the solution already exists
            QueryExpression queryCheckForSampleSolution = new QueryExpression
            {
                EntityName = Solution.EntityLogicalName,
                ColumnSet = new ColumnSet(),
                Criteria = new FilterExpression()
            };

            queryCheckForSampleSolution.Criteria.AddCondition("uniquename",
               ConditionOperator.Equal, solution.UniqueName);

            // Attempt to retrieve the solution
            EntityCollection querySampleSolutionResults =
               _serviceProxy.RetrieveMultiple(queryCheckForSampleSolution);

            // Create the solution if it doesn't already exist
            Solution SampleSolutionResults = null;

            if (querySampleSolutionResults.Entities.Count > 0)
            {
                SampleSolutionResults = (Solution)querySampleSolutionResults.Entities[0];
                _solutionsSampleSolutionId = (Guid)SampleSolutionResults.SolutionId;
            }

            if (SampleSolutionResults == null)
            {
                _solutionsSampleSolutionId = _serviceProxy.Create(solution);
            }
            */
        }
    }
}
