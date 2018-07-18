using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateSPSubsiteInCRM
{
    public class FolderStructure : IPlugin
    {
        #region Secure/Unsecure Configuration Setup
        private string _secureConfig = null;
        private string _unsecureConfig = null;

        public FolderStructure(string unsecureConfig, string secureConfig)
        {
            _secureConfig = secureConfig;
            _unsecureConfig = unsecureConfig;
        }
        #endregion
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);

            try
            {
                var entity = (Entity)context.InputParameters["Target"];

                var urlString = @"sharePointUrl";
                if (entity.LogicalName == "entityName" && entity.Contains("fieldName"))
                {
                    var rootRecord = entity.GetAttributeValue<string>("fieldName"); //Main or root record
                    if (CreateSharepointSite.CreateSubsite(service, rootRecord, urlString))
                        CreateSharePointSiteDocLocCRM(service, entity.Id, rootRecord, urlString);
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        public static void CreateSharePointSiteDocLocCRM(IOrganizationService service, Guid rootRecordId, string rootRecord, string siteUrl)
        {
            try
            {
                var newSiteUrl = siteUrl + "/" + rootRecord + "/";
                var spSite = new Entity("sharepointsite");
                spSite["name"] = rootRecord;
                spSite["description"] = rootRecord;
                spSite["absoluteurl"] = newSiteUrl;
                var _spSiteId = service.Create(spSite);

                if (_spSiteId != Guid.Empty)
                {
                    var spDocLoc = new Entity("sharepointdocumentlocation");
                    spDocLoc["name"] = rootRecord + " Root Document Library";
                    spDocLoc["description"] = rootRecord + " Root Document Library";
                    spDocLoc["parentsiteorlocation"] = new EntityReference("sharepointsite", _spSiteId);
                    spDocLoc["relativeurl"] = "sharepoinrUrl";
                    var spDocLocCreate = service.Create(spDocLoc);

                    if (spDocLocCreate != Guid.Empty)
                    {
                        var spDocLoc2 = new Entity("sharepointdocumentlocation");
                        spDocLoc2["name"] = rootRecord;
                        spDocLoc2["description"] = rootRecord + " Documents";
                        spDocLoc2["parentsiteorlocation"] = new EntityReference("sharepointdocumentlocation", spDocLocCreate);
                        spDocLoc2["relativeurl"] = "Documents";
                        spDocLoc2["regardingobjectid"] = new EntityReference("rootEntityName", rootRecordId);
                        var spDocLocCreate2 = service.Create(spDocLoc2);
                    }
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }
    }
}
