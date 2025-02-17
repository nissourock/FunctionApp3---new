using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using PnP.Framework;
using System.Globalization;
using System.Text;
using static FunctionApp3.Models;



namespace FunctionApp3
{
    public class AppSettings
    {
        public string AzureFunctionBaseURL { get; set; }
        // Add other configuration properties as needed.
    }
    public class WebhookResponder
    {

        private readonly ILogger<WebhookResponder> _logger;
        private string AzureFunctionBaseURL = Environment.GetEnvironmentVariable("AzureFunctionBaseURL")?.Trim().Replace("\"", "");

        public WebhookResponder(ILogger<WebhookResponder> logger) {
            _logger = logger;

        }
        [Function("WebhookResponder")]

        public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req) {
            _logger.LogInformation("SharePoint webhook received a request.");

            string validationToken = req.Query["validationtoken"];

            // If a validation token is present, we need to respond within 5 seconds by
            // returning the given validation token. This only happens when a new
            // webhook is being added or resubscribed every 6 months
            if (validationToken != null)
            {
                _logger.LogInformation($"Validation token {validationToken} received");
                return (ActionResult)new OkObjectResult(validationToken);
            }

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(requestBody).Value;

            if (requestBody != null)
            {
                // Handle the SharePoint webhook notification here
                _logger.LogInformation($"Webhook notification received: {requestBody}");

                // Example: Parse and process the notification data
                string? siteUrl = notifications.FirstOrDefault()?.SiteUrl;
                string? resourceId = notifications.FirstOrDefault()?.Resource;
                string? clientstate = notifications.FirstOrDefault()?.ClientState;


                _logger.LogInformation($"Site URL: {siteUrl}, Ressource ID: {resourceId}");
                // The warning here is warranted because we don't expect the response from the request. fire and forget (because we need to send an OK response in 5seconds for the webhook)
                if (clientstate == "ListWebhook")
                {
                    Task.Run(async () =>
                    {
                        using (var httpClient = new HttpClient())
                        {
                            var payload = new
                            {
                                siteUrl,
                                resourceId
                            };
                            var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");


                            await httpClient.PostAsync(AzureFunctionBaseURL + "/api/DossierMaitreCreator/", content);
                        }
                    });
                }
                if (clientstate == "CSVWebhook")
                {
                    Task.Run(async () =>
                    {
                        using (var httpClient = new HttpClient())
                        {
                            var payload = new
                            {
                                siteUrl,
                                resourceId
                            };
                            var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");


                            await httpClient.PostAsync(AzureFunctionBaseURL + "/api/CSVParser/", content);
                        }
                    });
                }

                return new OkResult();
            }



            return new OkResult();
        }



    }
    public class CSVParser
    {
        //PPR
        public string? siteURL = Environment.GetEnvironmentVariable("SharePoint_SiteUrl_GeostockPPR")?.Trim().Replace("\"", "");
        public string? AzureBlobStorageConnectionString = Environment.GetEnvironmentVariable("AzureBlobStorageConnectionStringPPR")?.Trim().Replace("\"", "");
        public string? CSVDocumentLibraryTitle = Environment.GetEnvironmentVariable("CSVDocumentLibraryTitlePPR")?.Trim().Replace("\"", "");

        //PROD
        //public string? siteURL = Environment.GetEnvironmentVariable("SharePoint_SiteUrl_GeostockPROD")?.Trim().Replace("\"", "");
        //public string? AzureBlobStorageConnectionString = Environment.GetEnvironmentVariable("AzureBlobStorageConnectionString")?.Trim().Replace("\"", "");
        //public string? CSVDocumentLibraryTitle = Environment.GetEnvironmentVariable("CSVDocumentLibraryTitle")?.Trim().Replace("\"", "");


        //Indiscriminate PPR/PROD
        public string? clientID = Environment.GetEnvironmentVariable("SharePoint_ClientID")?.Trim().Replace("\"", "");
        public string? clientSecretID = Environment.GetEnvironmentVariable("SharePoint_ClientSecretID")?.Trim().Replace("\"", "");
        public string? BlobContainerCSV = Environment.GetEnvironmentVariable("BlobContainerCSV")?.Trim().Replace("\"", "");
        public string? BlobContainerList = Environment.GetEnvironmentVariable("BlobContainerList")?.Trim().Replace("\"", "");
        public string? CreationListName = Environment.GetEnvironmentVariable("CreationListName")?.Trim().Replace("\"", "");

        private readonly ILogger<CSVParser> _logger;
        public static string EscapeSingleQuotes(string input) {
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }

            return input.Replace("'", "\\'").Replace("\"", "\\\"");
        }
        public CSVParser(ILogger<CSVParser> logger) {
            _logger = logger;

        }
        [Function("CSVParser")]

        public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req) {
            _logger.LogInformation("CSV" +
                "" +
                " webhook received a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic? data = JsonConvert.DeserializeObject(requestBody);

            // Extract parameters from the request body
            string? siteUrl = data?.siteUrl;
            string? resourceId = data?.resourceId;
            using (var clientContext = new AuthenticationManager().GetACSAppOnlyContext(siteURL, clientID, clientSecretID))
            {
                using (PnPContext pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(clientContext))
                {

                    var changeQuery = new ChangeQueryOptions(false, false)
                    {
                        Item = true,
                        FetchLimit = 999,
                        File = true,
                        Folder = true,
                        Add = true,
                        Rename = true,
                    };


                    var lastChangeToken = await Services.GetLatestChangeTokenAsync(resourceId, AzureBlobStorageConnectionString, BlobContainerCSV, _logger);

                    if (lastChangeToken != null && !String.IsNullOrEmpty(lastChangeToken))
                    {
                        changeQuery.ChangeTokenStart = new ChangeTokenOptions(lastChangeToken);
                    }

                    var targetList = pnpCoreContext.Web.Lists.GetById(Guid.Parse(resourceId), p => p.Title,
                                                                                p => p.Fields.QueryProperties(p => p.InternalName,
                                                                                p => p.PrimaryFieldId,
                                                                                p => p.FieldTypeKind,
                                                                                p => p.TypeAsString,
                                                                                p => p.Title, p => p.All));
                    _logger.LogInformation($"SharePoint List Name: {targetList.Title}");

                    var changes = await targetList.GetChangesAsync(changeQuery);

                    if (changes.Any())
                    {
                        Services.SaveLatestChangeTokenAsync(changes.Last().ChangeToken, resourceId, AzureBlobStorageConnectionString, BlobContainerCSV, _logger);
                    }
                    var addChangesList = changes.Where(change => change.ChangeType == PnP.Core.Model.SharePoint.ChangeType.Add).ToList();
                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers);

                    if (addChangesList.Count > 0)
                    {

                        foreach (var change in addChangesList)
                        {
                            if (change is IChangeItem changeItem)
                            {
                                if (changeItem.IsPropertyAvailable<IChangeItem>(i => i.ItemId))
                                {
                                    string operationIdentifierGUID = Guid.NewGuid().ToString();
                                    var itemId = changeItem.ItemId;

                                    var item = await targetList.Items.GetByIdAsync(itemId, i => i.Title, i => i.All, i => i.File);
                                    if (item != null)
                                    {
                                        if (item.FileSystemObjectType == PnP.Core.Model.SharePoint.FileSystemObjectType.File)
                                        {
                                            var authorField = (item["Author"] as IFieldUserValue);
                                            var authorMail = pnpCoreContext.Web.SiteUsers.AsRequested().Where(i => i.Id == authorField?.LookupId).FirstOrDefault()?.Mail;
                                            var serverRelativeURL = item.File.ServerRelativeUrl;
                                            if (serverRelativeURL.ToString().ToUpperInvariant().Contains("CSV"))
                                            {

                                                var CSVfile = await pnpCoreContext.Web.GetFileByServerRelativeUrlAsync(serverRelativeURL);

                                                using var stream = await CSVfile.GetContentAsync();
                                                using var reader = new StreamReader(stream, Encoding.Latin1); // Ensure Latin encoding because fields may contain accents (french)
                                                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                                                {
                                                    HeaderValidated = null,
                                                    MissingFieldFound = null,
                                                    BadDataFound = context =>
                                                    {
                                                        Console.WriteLine($"Bad data found on row : {context.RawRecord}");
                                                    },
                                                    Delimiter = ";",
                                                    TrimOptions = TrimOptions.Trim
                                                };
                                                var csv = new CsvReader(reader, config);


                                                var records = csv.GetRecords<dynamic>().ToList();
                                                foreach (var record in records)
                                                {
                                                    var SiteURL = record.SiteURL;
                                                    var Type = record.Type;
                                                    var DossierPays = record.DossierPays;
                                                    var Statut = record.Statut;
                                                    var Titre = record.Titre;
                                                    var DossierProjet = record.DossierProjet;


                                                    //persons
                                                    //var emetteursList = record.emetteursList;
                                                    var emetteursList = record.emetteursList;
                                                    var verificateursList = record.verificateursList;
                                                    var approbateursList = record.approbateurs;
                                                    var chef = record.ChefDeProjet;
                                                    var secretaire = record.secretaire;
                                                    var intervenantPAO = record.intervenantPAO;
                                                    var intervenantTTX = record.intervenantTTX;

                                                    //initialize processed variables
                                                    var emetteursListProcessed = "";
                                                    var verificateursListProcessed = "";
                                                    var approbateursListProcessed = "";
                                                    var chefProcessed = "";
                                                    var secretaireProcessed = "";
                                                    var intervenantPAOProcessed = "";
                                                    var intervenantTTXProcessed = "";
                                                    var prefix = "i:0#.f|membership|";
                                                    //traitement des personnes 
                                                    if (!string.IsNullOrEmpty(emetteursList))
                                                    {
                                                        emetteursListProcessed = Services.FormatEmails(emetteursList, prefix, _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(verificateursList))
                                                    {
                                                        verificateursListProcessed = Services.FormatEmails(verificateursList, prefix, _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(approbateursList))
                                                    {
                                                        approbateursListProcessed = Services.FormatEmails(approbateursList, prefix, _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(chef))
                                                    {
                                                        chefProcessed = Services.FormatEmails(chef, prefix, _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(secretaire))
                                                    {
                                                        secretaireProcessed = Services.FormatEmails(secretaire, prefix, _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(intervenantPAOProcessed))
                                                    {
                                                        intervenantPAOProcessed = Services.FormatEmails(intervenantPAOProcessed, prefix, _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(intervenantTTXProcessed))
                                                    {
                                                        intervenantTTXProcessed = Services.FormatEmails(intervenantTTXProcessed, prefix, _logger);
                                                    }




                                                    //lookups
                                                    var client = record.client;
                                                    var sites = record.sites;
                                                    var motsCles = record.motsCles;
                                                    var projet = record.Projet;

                                                    var societe = record.societe;
                                                    var departement = record.departement;

                                                    //metadonnées gérées
                                                    var specialiteMetier = EscapeSingleQuotes(record.specialiteMetier);
                                                    var typeDeDocument = EscapeSingleQuotes(record.typeDeDocument);

                                                    var Description = EscapeSingleQuotes(record.Description);
                                                    var activite = EscapeSingleQuotes(record.activite);
                                                    var bibliothequeCible = EscapeSingleQuotes(record.bibliothequeCible);

                                                    var langue = EscapeSingleQuotes(record.langue);
                                                    var phaseEtude = EscapeSingleQuotes(record.phaseEtude);

                                                    var nombrePage = record.nombrePage;
                                                    var nombreAnnexes = record.nombreAnnexes;
                                                    var DateSouhaitee = record.DateSouhaitee;
                                                    var Etat = EscapeSingleQuotes(record.Etat);

                                                    // traitement des lookups

                                                    //initialisation des variables vides 
                                                    var projetID = "";
                                                    var paysID = "";
                                                    var sitesID = "";
                                                    var motsClesIDs = "";
                                                    var societeIDs = "";
                                                    var departementIDs = "";
                                                    var clientIDs = "";

                                                    if (!string.IsNullOrEmpty(client))
                                                    {
                                                        clientIDs = Services.GetItemIdsBySiteTitles(clientContext, "Clients", client, "Nom_x0020_Client",  _logger);
                                                    }

                                                    if (!string.IsNullOrEmpty(projet))
                                                    {
                                                        projetID = Services.GetItemIdsBySiteTitles(clientContext, "Codes Projets", projet, "Libell_x00e9_", _logger);
                                                    }


                                                    if (!string.IsNullOrEmpty(DossierPays))
                                                    {
                                                        paysID = Services.GetItemIdsBySiteTitles(clientContext, "Codes Pays", DossierPays, "Libell_x00e9_", _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(sites))
                                                    {
                                                        sitesID = Services.GetItemIdsBySiteTitles(clientContext, "Liste des Sites", sites, "Site", _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(motsCles))
                                                    {
                                                        motsClesIDs = Services.GetItemIdsBySiteTitles(clientContext, "Mots-Clés", motsCles, "Title", _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(societe))
                                                    {
                                                        societeIDs = Services.GetItemIdsBySiteTitles(clientContext, "Société", societe, "Title", _logger);
                                                    }
                                                    if (!string.IsNullOrEmpty(departement))
                                                    {
                                                        departementIDs = Services.GetItemIdsBySiteTitles(clientContext, "Departements", departement, "Libell_x00e9_", _logger);
                                                    }




                                                    string JSONTemplate = $@"{{'ID':'placeholder','Title':'{Titre}','CodeProjet':'{projetID}','Source':'','NumeroDossier':'','Activite':'{activite}','Approbateurs':'{approbateursListProcessed}','Emetteurs':'{emetteursListProcessed}','Verificateurs':'{verificateursListProcessed}','Secretaire':'{secretaireProcessed}','Chef':'{chefProcessed}','Bibliotheque':'{bibliothequeCible}','Langue':'{langue}','DateSouhaite':'{DateSouhaitee}','NombrePage':'{nombrePage}','NombreAnnexes':'{nombreAnnexes}','Departement':'{departementIDs}','DossierProjet':'{DossierProjet}','Client':'{clientIDs}','Etat':'{Etat}','IntervenantPAO':'{intervenantPAOProcessed}','IntervenantTTX':'{intervenantTTXProcessed}','PhaseEtude':'{phaseEtude}','SpecialiteMetier':'{specialiteMetier}','TypeDeDocument':'{typeDeDocument}','Societe':'{societeIDs}','Revision':'','ObjetRevision':'','CodePays':'{paysID}','IntituleProjet':'','MotsCles':'{motsClesIDs}','Description':'{Description}','Site':'{sitesID}'}}";
                                                    string JSONTemplateQuotes = $"\"{JSONTemplate}\"";
                                                    string JSON = "\"" + JSONTemplate + "\"";
                                                    //add the list item
                                                    var webhookCreateList = pnpCoreContext.Web.Lists.GetByTitle("Dossiers maitres création");
                                                    Dictionary<string, object> listItem = new Dictionary<string, object>()
                        {
                            { "FolderPathDestination", ""+DossierPays+"/" },
                            {"creatorEmail",authorMail!= null? authorMail : "powerauto_p2022121914@vinci-construction.com" },
                            { "SiteURL", ""+SiteURL },
                            { "ListName", ""+Type },
                            { "JSON", "\""+JSONTemplate+"\"" },
                            { "Statutwebhook", ""+"En attente de création"+"" },
                            {"ModeCreation", "Multi" },
                            {"OperationIdentifier",operationIdentifierGUID }


                        };
                                                    var addedItem = await webhookCreateList.Items.AddAsync(listItem);
                                                }
                                            }
                                        };
                                        ;
                                    }


                                }
                            };
                        }

                    }
                }
                return new OkResult();
            }
        }

    }
    public class DossierMaitreCreator
    {
        //PPR
        public string? siteURL = Environment.GetEnvironmentVariable("SharePoint_SiteUrl_GeostockPPR")?.Trim().Replace("\"", "");
        public string? AzureBlobStorageConnectionString = Environment.GetEnvironmentVariable("AzureBlobStorageConnectionStringPPR")?.Trim().Replace("\"", "");
        public string? CSVDocumentLibraryTitle = Environment.GetEnvironmentVariable("CSVDocumentLibraryTitlePPR")?.Trim().Replace("\"", "");

        //PROD
        //public string? siteURL = Environment.GetEnvironmentVariable("SharePoint_SiteUrl_GeostockPROD")?.Trim().Replace("\"", "");
        //public string? AzureBlobStorageConnectionString = Environment.GetEnvironmentVariable("AzureBlobStorageConnectionString")?.Trim().Replace("\"", "");
        //public string? CSVDocumentLibraryTitle = Environment.GetEnvironmentVariable("CSVDocumentLibraryTitle")?.Trim().Replace("\"", "");


        //Indiscriminate PPR/PROD
        public string? clientID = Environment.GetEnvironmentVariable("SharePoint_ClientID")?.Trim().Replace("\"", "");
        public string? clientSecretID = Environment.GetEnvironmentVariable("SharePoint_ClientSecretID")?.Trim().Replace("\"", "");
        public string? BlobContainerCSV = Environment.GetEnvironmentVariable("BlobContainerCSV")?.Trim().Replace("\"", "");
        public string? BlobContainerList = Environment.GetEnvironmentVariable("BlobContainerList")?.Trim().Replace("\"", "");
        public string? CreationListName = Environment.GetEnvironmentVariable("CreationListName")?.Trim().Replace("\"", "");

        private readonly ILogger<DossierMaitreCreator> _logger;

        public DossierMaitreCreator(ILogger<DossierMaitreCreator> logger) {
            _logger = logger;
        }

        [Function("DossierMaitreCreator")]

        public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req) {
            _logger.LogInformation("Dossier maitre creator" +
                " webhook received a request.");
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic? data = JsonConvert.DeserializeObject(requestBody);

           

            // Extract parameters from the request body
            string? siteUrl = data?.siteUrl;
            _logger.LogInformation($"site URL: {siteUrl}");

            string? resourceId = data?.resourceId;
            _logger.LogInformation($"Ressource ID: {resourceId}");
            using (var clientContext = new AuthenticationManager().GetACSAppOnlyContext(siteURL, clientID, clientSecretID))
            {
                using (PnPContext pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(clientContext))
                {
                    var targetList = pnpCoreContext.Web.Lists.GetById(Guid.Parse(resourceId), p => p.Title,
                                                                                p => p.Fields.QueryProperties(p => p.InternalName,
                                                                                p => p.PrimaryFieldId,
                                                                                p => p.FieldTypeKind,
                                                                                p => p.TypeAsString,
                                                                                p => p.Title, p => p.All));




                    List<dynamic?> results = new List<dynamic?>();

                    //Use this for faster testing 
                    // var result =  await Services.DossierMaitreCreation(clientContext, pnpCoreContext, 13, targetList, _logger);
                    //var result2 = await Services.DossierMaitreCreation(clientContext, pnpCoreContext, 268, targetList, _logger);
                    //var result3 = await Services.DossierMaitreCreation(clientContext, pnpCoreContext, 262, targetList, _logger);



                     //var result4 = await Services.DossierMaitreCreation(clientContext, pnpCoreContext, 743, targetList, _logger);



                    //  results.Add(result4);
                    //results.Add(result2);
                    //results.Add(result3);
                    //results.Add(result4);
                    //        var groupedItems = results
                    //.GroupBy(item => item.OperationIdentifier)
                    //.Select(group => {
                    //    Console.WriteLine($"Grouping {group.Key}: Count = {group.Count()}");
                    //    return group.Count() > 1
                    //        ? (object)group.ToList()
                    //        : (object)group.First();
                    //})
                    //.ToList();

                    //        var singleItems = results?
                    //.Where(item => string.IsNullOrEmpty(item?.OperationIdentifier))
                    //.Select(item => (object)item)
                    //.ToList();

                    //        // Group remaining items by OperationIdentifier
                    //        var groupedItems = results?
                    //            .Where(item => !string.IsNullOrEmpty(item?.OperationIdentifier))
                    //            .GroupBy(item => item?.OperationIdentifier)
                    //            .Select(group => group.Count() > 1
                    //                ? (object)group.ToList() // Group with more than one item
                    //                : (object)new List<dynamic> { group.First() }) // Single item in a list
                    //        .ToList();

                    //        try { Services.Mailer2(clientContext,pnpCoreContext,groupedItems,singleItems); }
                    //        catch (Exception ex)
                    //        {
                    //            Console.WriteLine("Email send unsuccessful");
                    //            Console.WriteLine(ex.Message);

                    //        }




                    var changeQuery = new ChangeQueryOptions(false, false)
                    {
                        Item = true,
                        FetchLimit = 999,
                        File = true,
                        Folder = true,
                        Add = true,

                    };


                    var lastChangeToken = await Services.GetLatestChangeTokenAsync(resourceId, AzureBlobStorageConnectionString, BlobContainerList, _logger);

                    if (lastChangeToken != null && !String.IsNullOrEmpty(lastChangeToken))
                    {
                        changeQuery.ChangeTokenStart = new ChangeTokenOptions(lastChangeToken);
                    }


                    _logger.LogInformation($"SharePoint List Name: {targetList.Title}");

                    var changes = await targetList.GetChangesAsync(changeQuery);

                    if (changes.Any())
                    {
                        Services.SaveLatestChangeTokenAsync(changes.Last().ChangeToken, resourceId, AzureBlobStorageConnectionString, BlobContainerList, _logger);
                    }
                    var addChangesList = changes.Where(change => change.ChangeType == PnP.Core.Model.SharePoint.ChangeType.Add).ToList();
                    if (addChangesList.Count > 0)
                    {
                        foreach (var change in addChangesList)
                        {
                            if (change is IChangeItem changeItem)
                            {
                                if (changeItem.IsPropertyAvailable<IChangeItem>(i => i.ItemId))
                                {
                                    var resultAsync = await Services.DossierMaitreCreation(clientContext, pnpCoreContext, changeItem.ItemId, targetList, _logger);
                                    results.Add(resultAsync);

                                }
                            }
                        }
                        var singleItemsAsync = results?
            .Where(item => string.IsNullOrEmpty(item?.OperationIdentifier))
            .Select(item => (object?)item)
            .ToList();

                        // Group remaining items by OperationIdentifier
                        var groupedItemsAsync = results?
                        .Where(item => !string.IsNullOrEmpty(item?.OperationIdentifier))
                        .GroupBy(item => item?.OperationIdentifier)
                        .Select(group => group.Count() > 1
                            ? (object)group.ToList() // Group with more than one item
                            : (object)new List<dynamic> { group.First() }) // Single item in a list
                    .ToList();
                        try { Services.Mailer2(clientContext, pnpCoreContext, groupedItemsAsync, singleItemsAsync, _logger); }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Email send unsuccessful");
                            Console.WriteLine(ex.Message);

                        }


                    }
                }

                return new OkResult();
            }



        }


    }
}


