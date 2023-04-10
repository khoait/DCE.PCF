using DCE.PCF.Plugins.Modals;
using DCE.PCF.Plugins.Utils;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DCE.PCF.Plugins
{
    [CrmPluginRegistration("Create",
    "storm_course", StageEnum.PostOperation, ExecutionModeEnum.Synchronous,
    "", "PostCreate_storm_course_Sync", 1,
    IsolationModeEnum.Sandbox
    , UnSecureConfiguration = "storm_courses_coursecategories,storm_selectedcategories"
    , Image1Type = ImageTypeEnum.PostImage
    , Image1Name = "PostImage"
    , Image1Attributes = "storm_selectedcategories"
    , Description = "Associate selected categories from PolyLookup"
    )]
    public class AssociatePolyLookup : IPlugin
    {
        public string RelationshipName { get; set; }
        public string SelectedItemsField { get; set; }
        public AssociatePolyLookup(string unsecureString, string secureString)
        {
            var strings = unsecureString.Split(',');
            if (strings.Length == 2)
            {
                RelationshipName = strings[0].Trim();
                SelectedItemsField = strings[1].Trim();
            }
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            var executionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var organizationServiceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var currentUserService = organizationServiceFactory.CreateOrganizationService(executionContext.UserId);
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            var target = (Entity)executionContext.InputParameters["Target"];
            if (target == null)
            {
                throw new InvalidPluginExecutionException(OperationStatus.Failed, 1, "Target is null");
            }

            var postImage = executionContext.PostEntityImages["PostImage"];
            if (postImage == null)
            {
                throw new InvalidPluginExecutionException(OperationStatus.Failed, 1, "Post Image is null");
            }

            if (string.IsNullOrEmpty(RelationshipName) || string.IsNullOrEmpty(SelectedItemsField))
            {
                throw new InvalidPluginExecutionException(OperationStatus.Failed, 1, "Relationship Name or Selected Items Field is null");
            }

            tracingService.Trace("Relationship name: {0}", RelationshipName);
            tracingService.Trace("Selected items field: {0}", SelectedItemsField);

            var selectedCategoriesText = postImage.GetAttributeValue<string>(SelectedItemsField);

            tracingService.Trace("Selected categories text: {0}", selectedCategoriesText);

            if (string.IsNullOrEmpty(selectedCategoriesText)) return;

            try
            {
                var selectedItems = JsonConvert.Deserialize<List<PolyLookupItem>>(selectedCategoriesText);
                if (!selectedItems.Any()) return;


                var selectedRefs = selectedItems.Select(i => new EntityReference(i.Etn, Guid.Parse(i.Id))).ToList();

                tracingService.Trace("Selected Refs: {0}", selectedRefs);

                currentUserService.Associate(target.LogicalName, target.Id, new Relationship(RelationshipName), new EntityReferenceCollection(selectedRefs));

                tracingService.Trace("Done");
            }
            catch (Exception)
            {

                throw new InvalidPluginExecutionException(OperationStatus.Failed, 1, "Failed to associate selected items");
            }

        }
    }
}
