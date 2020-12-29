using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;

namespace NDXC.Blog.CWA
{
    public sealed class GetEnvironmentVariableValue : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {

            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            tracingService.Trace("Setting up...");

            //Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("Querying Dataverse Environment Variable Definitions...");

            // Only 1 Definition is expected 
            EntityCollection ecDefinitions = GetDefinitions(executionContext, service);

            if (ecDefinitions.Entities.Count == 1)
            {

                tracingService.Trace("Environment Variable Definition Found...");
                tracingService.Trace("Setting Default Value...");

                EnvironmentVariableValue.Set(executionContext, (string)ecDefinitions.Entities[0]["defaultvalue"]);

                tracingService.Trace("Querying Dataverse Environment Variable Values for Definition...");

                // 0, or 1 Value is expected
                // 0 if no Current Value has been set for the Environment Variable
                EntityCollection ecValues = GetValues(executionContext, service, ecDefinitions); ;

                if (ecValues.Entities.Count == 1)
                {

                    tracingService.Trace("Environment Variable Value Found...");
                    tracingService.Trace("Setting Value...");

                    EnvironmentVariableValue.Set(executionContext, (string)ecValues.Entities[0]["value"]);
                }
            }
            else
            {

                throw new InvalidPluginExecutionException(string.Format("Environment Variable with Name {0} does not exist", EnvironmentVariableName));
            }
        }

        private EntityCollection GetDefinitions(CodeActivityContext executionContext, IOrganizationService service)
        {

            var qeDefinitions = new QueryExpression();
            qeDefinitions.EntityName = "environmentvariabledefinition";
            qeDefinitions.ColumnSet = new ColumnSet();
            qeDefinitions.ColumnSet.Columns.Add("environmentvariabledefinitionid");
            qeDefinitions.ColumnSet.Columns.Add("defaultvalue");
            qeDefinitions.TopCount = 1;

            qeDefinitions.Criteria.AddFilter(LogicalOperator.And);
            qeDefinitions.Criteria.AddCondition("displayname", ConditionOperator.Equal, EnvironmentVariableName.Get<string>(executionContext));

            return service.RetrieveMultiple(qeDefinitions);
        }

        private EntityCollection GetValues(CodeActivityContext executionContext, IOrganizationService service, EntityCollection ecDefinitions)
        {

            var qeValues = new QueryExpression();
            qeValues.EntityName = "environmentvariablevalue";
            qeValues.ColumnSet = new ColumnSet();
            qeValues.ColumnSet.Columns.Add("value");
            qeValues.TopCount = 1;

            qeValues.Criteria.AddFilter(LogicalOperator.And);
            qeValues.Criteria.AddCondition("environmentvariabledefinitionid", ConditionOperator.Equal, ecDefinitions.Entities[0].Id);

            return service.RetrieveMultiple(qeValues);
        }

        [Input("Environment Variable Name")]
        public InArgument<string> EnvironmentVariableName { get; set; }

        [Output("Environment Variable Value")]
        public OutArgument<string> EnvironmentVariableValue { get; set; }
    }
}
