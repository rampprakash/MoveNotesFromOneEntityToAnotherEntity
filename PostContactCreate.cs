using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;


namespace Dataverse.MoveNotesFromOnetoAnother
{
    public class PostContactCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service

            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider. 
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            // Obtain the organization service reference which you will need for 
            // web service calls. 
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                Entity getContactDetails = (Entity)context.InputParameters["Target"];
                string getNotes = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
  <entity name='annotation'>
    <all-attributes />
    <order attribute='subject' descending='false' />
    <link-entity name='lead' from='leadid' to='objectid' link-type='inner' alias='aa'>
      <filter type='and'>
        <condition attribute='leadid' operator='eq' value='{objectid}' />
      </filter>
    </link-entity>
  </entity>
</fetch>";

                getNotes = getNotes.Replace("{objectid}", "LEADGUID");
                EntityCollection collectNotes = service.RetrieveMultiple(new FetchExpression(getNotes));

                foreach (var loopColelctNotes in collectNotes.Entities)
                {
                    Entity _annotation = new Entity("annotation");
                    _annotation.Attributes["objectid"] = new EntityReference("contact", getContactDetails.Id);
                    _annotation.Attributes["objecttypecode"] = "contact";
                    if (loopColelctNotes.Attributes.Contains("subject"))
                    {
                        _annotation.Attributes["subject"] = loopColelctNotes.Attributes["subject"];
                    }

                    if (loopColelctNotes.Attributes.Contains("documentbody"))
                    {
                        _annotation.Attributes["documentbody"] = loopColelctNotes.Attributes["documentbody"];
                    }

                    if (loopColelctNotes.Attributes.Contains("mimetype"))
                    {
                        _annotation.Attributes["mimetype"] = loopColelctNotes.Attributes["mimetype"];
                    }

                    if (loopColelctNotes.Attributes.Contains("notetext"))
                    {
                        _annotation.Attributes["notetext"] = loopColelctNotes.Attributes["notetext"];
                    }

                    if (loopColelctNotes.Attributes.Contains("filename"))
                    {
                        _annotation.Attributes["filename"] = loopColelctNotes.Attributes["filename"];
                    }

                    service.Create(_annotation);

                }
            }
            catch (InvalidPluginExecutionException ex)
            {

                throw new InvalidPluginExecutionException(ex.Message.ToString());
            }
        }
    }
}