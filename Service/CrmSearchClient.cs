using Microsoft.Crm.Sdk.Messages;
using Microsoft.IdentityModel.Protocols;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace CrmKbBot.Service
{
    public class CrmSearchClient
    {
        private readonly IOrganizationService orgService;

        public CrmSearchClient()
        {
            var connectionString = ConfigurationManager.ConnectionStrings["CRM"].ConnectionString;

            CrmServiceClient conn = new CrmServiceClient(connectionString);
            this.orgService = (IOrganizationService)conn.OrganizationServiceProxy;
        }

        public async Task<FullTextSearchKnowledgeArticleResponse> SearchAsync(string searchText)
        {
            var queryExpression = new QueryExpression("knowledgearticle");
            queryExpression.ColumnSet = new ColumnSet(true);
            queryExpression.PageInfo = new PagingInfo();
            queryExpression.PageInfo.PageNumber = 1;
            queryExpression.PageInfo.Count = 50;

            var request = new FullTextSearchKnowledgeArticleRequest();
            request.RemoveDuplicates = true;
            request.StateCode = 3;
            request.UseInflection = true;
            request.SearchText = searchText;
            request.QueryExpression = queryExpression;

            return (FullTextSearchKnowledgeArticleResponse)this.orgService.Execute(request);
        }

        public async Task<string> CreateCase(string description)
        {
            var accountFetch = "<fetch count='1'><entity name='account'><attribute name='accountid'/></entity></fetch>";
            var account = orgService.RetrieveMultiple(new FetchExpression(accountFetch)).Entities.First();

            var caseEntity = new Entity("incident");
            caseEntity["title"] = description;
            caseEntity["customerid"] = new EntityReference("account", account.Id);

            var caseId = orgService.Create(caseEntity);

            var updatedCase = this.orgService.Retrieve("incident", caseId, new ColumnSet("ticketnumber"));
            return (string)updatedCase["ticketnumber"];
        }
    }
}