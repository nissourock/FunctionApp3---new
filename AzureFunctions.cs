using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
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
    
    public class WebhookResponder
    {

        private readonly ILogger<WebhookResponder> _logger;
        private readonly FunctionAppSettings _settings;
        private readonly IHttpClientFactory _httpClientFactory;


        public WebhookResponder(IOptions<FunctionAppSettings> options, IHttpClientFactory httpClientFactory, ILogger<WebhookResponder> logger) {
            _logger = logger;
            _settings = options.Value;
            _httpClientFactory = httpClientFactory;
        }
        [Function("WebhookResponder")]

        public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req) {
            _logger.LogInformation("SharePoint webhook received a request.");

            string? validationToken = req.Query["validationtoken"];

            // If a validation token is present, we need to respond within 5 seconds by
            // returning the given validation token. This only happens when a new
            // webhook is being added or resubscribed every 6 months
            if (validationToken != null)
            {
                _logger.LogInformation($"Validation token {validationToken} received");
                return (ActionResult) new OkObjectResult(validationToken);
            }

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(requestBody)?.Value;

            if (requestBody != null)
            {
                // Handle the SharePoint webhook notification here
                _logger.LogInformation($"Webhook notification received: {requestBody}");

                // Example: Parse and process the notification data
                string? siteUrl = notifications?.FirstOrDefault()?.SiteUrl;
                string? resourceId = notifications?.FirstOrDefault()?.Resource;
                string? clientstate = notifications?.FirstOrDefault()?.ClientState;


                _logger.LogInformation($"Site URL: {siteUrl}, Ressource ID: {resourceId}");
                // The warning here is warranted because we don't expect the response from the request. fire and forget (because we need to send an OK response in 5seconds for the webhook)
                if (clientstate == "ListWebhook")
                {
                    Task.Run(async () =>
                    {
                        using (var httpClient = _httpClientFactory.CreateClient())
                        {
                            var payload = new
                            {
                                siteUrl,
                                resourceId
                            };
                            var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");


                            await httpClient.PostAsync(_settings.AzureFunctionBaseURL + "/api/DossierMaitreCreator/", content);
                        }
                    });
                }
                if (clientstate == "CSVWebhook")
                {
                    Task.Run(async () =>
                    {
                        using (var httpClient = _httpClientFactory.CreateClient())
                        {
                            var payload = new
                            {
                                siteUrl,
                                resourceId
                            };
                            var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");


                            await httpClient.PostAsync(_settings.AzureFunctionBaseURL + "/api/CSVParser/", content);
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
       

        private readonly ILogger<CSVParser> _logger;
        private readonly FunctionAppSettings _settings;
        private readonly ISharePointContextFactory _spContextFactory;
        private readonly IServices _services;

        public CSVParser(IOptions<FunctionAppSettings> options, ISharePointContextFactory spContextFactory, ILogger<CSVParser> logger, IServices services) {
            _settings = options.Value;
            _logger = logger;
            _spContextFactory = spContextFactory;
            _services = services;
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
            using (var clientContext = _spContextFactory.CreateClientContext(_settings.siteURL, _settings.clientID, _settings.clientSecretID))
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


                    var lastChangeToken = await _services.GetLatestChangeTokenAsync(resourceId, _settings.AzureBlobStorageConnectionString, _settings.BlobContainerCSV);

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
                        await _services.SaveLatestChangeTokenAsync(changes.Last().ChangeToken, resourceId, _settings.AzureBlobStorageConnectionString, _settings.BlobContainerCSV);
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
                                                        emetteursListProcessed = _services.FormatEmails(emetteursList, prefix);
                                                    }
                                                    if (!string.IsNullOrEmpty(verificateursList))
                                                    {
                                                        verificateursListProcessed = _services.FormatEmails(verificateursList, prefix);
                                                    }
                                                    if (!string.IsNullOrEmpty(approbateursList))
                                                    {
                                                        approbateursListProcessed = _services.FormatEmails(approbateursList, prefix);
                                                    }
                                                    if (!string.IsNullOrEmpty(chef))
                                                    {
                                                        chefProcessed = _services.FormatEmails(chef, prefix);
                                                    }
                                                    if (!string.IsNullOrEmpty(secretaire))
                                                    {
                                                        secretaireProcessed = _services.FormatEmails(secretaire, prefix);
                                                    }
                                                    if (!string.IsNullOrEmpty(intervenantPAOProcessed))
                                                    {
                                                        intervenantPAOProcessed = _services.FormatEmails(intervenantPAOProcessed, prefix);
                                                    }
                                                    if (!string.IsNullOrEmpty(intervenantTTXProcessed))
                                                    {
                                                        intervenantTTXProcessed = _services.FormatEmails(intervenantTTXProcessed, prefix);
                                                    }




                                                    //lookups
                                                    var client = record.client;
                                                    var sites = record.sites;
                                                    var motsCles = record.motsCles;
                                                    var projet = record.Projet;

                                                    var societe = record.societe;
                                                    var departement = record.departement;

                                                    //metadonnées gérées
                                                    var specialiteMetier = _services.EscapeSingleQuotes(record.specialiteMetier);
                                                    var typeDeDocument = _services.EscapeSingleQuotes(record.typeDeDocument);

                                                    var Description = _services.EscapeSingleQuotes(record.Description);
                                                    var activite = _services.EscapeSingleQuotes(record.activite);
                                                    var bibliothequeCible = _services.EscapeSingleQuotes(record.bibliothequeCible);

                                                    var langue = _services.EscapeSingleQuotes(record.langue);
                                                    var phaseEtude = _services.EscapeSingleQuotes(record.phaseEtude);

                                                    var nombrePage = record.nombrePage;
                                                    var nombreAnnexes = record.nombreAnnexes;
                                                    var DateSouhaitee = record.DateSouhaitee;
                                                    var Etat = _services.EscapeSingleQuotes(record.Etat);

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
                                                        clientIDs = _services.GetItemIdsBySiteTitles(clientContext, "Clients", client, "Nom_x0020_Client");
                                                    }

                                                    if (!string.IsNullOrEmpty(projet))
                                                    {
                                                        projetID = _services.GetItemIdsBySiteTitles(clientContext, "Codes Projets", projet, "Libell_x00e9_");
                                                    }


                                                    if (!string.IsNullOrEmpty(DossierPays))
                                                    {
                                                        paysID = _services.GetItemIdsBySiteTitles(clientContext, "Codes Pays", DossierPays, "Libell_x00e9_");
                                                    }
                                                    if (!string.IsNullOrEmpty(sites))
                                                    {
                                                        sitesID = _services.GetItemIdsBySiteTitles(clientContext, "Liste des Sites", sites, "Site");
                                                    }
                                                    if (!string.IsNullOrEmpty(motsCles))
                                                    {
                                                        motsClesIDs = _services.GetItemIdsBySiteTitles(clientContext, "Mots-Clés", motsCles, "Title");
                                                    }
                                                    if (!string.IsNullOrEmpty(societe))
                                                    {
                                                        societeIDs = _services.GetItemIdsBySiteTitles(clientContext, "Société", societe, "Title");
                                                    }
                                                    if (!string.IsNullOrEmpty(departement))
                                                    {
                                                        departementIDs = _services.GetItemIdsBySiteTitles(clientContext, "Departements", departement, "Libell_x00e9_");
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
        private readonly ILogger<DossierMaitreCreator> _logger;
        private readonly FunctionAppSettings _settings;
        private readonly ISharePointContextFactory _spContextFactory;
        private readonly IServices _services;

        public DossierMaitreCreator(IOptions<FunctionAppSettings> options, ISharePointContextFactory spContextFactory, ILogger<DossierMaitreCreator> logger, IServices services) {
            _settings = options.Value;
            _logger = logger;
            _spContextFactory = spContextFactory;
            _services = services;
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
            using (var clientContext = _spContextFactory.CreateClientContext(_settings.siteURL, _settings.clientID, _settings.clientSecretID))
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

                    //    var result4 = await _services.DossierMaitreCreation(clientContext, pnpCoreContext,522, targetList, _settings.deploymentEnv);



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


                    _logger.LogInformation($"connected successfully");

                    var changeQuery = new ChangeQueryOptions(false, false)
                    {
                        Item = true,
                        FetchLimit = 999,
                        File = true,
                        Folder = true,
                        Add = true,

                    };


                    var lastChangeToken = await _services.GetLatestChangeTokenAsync(resourceId, _settings.AzureBlobStorageConnectionString, _settings.BlobContainerList);

                    //var parts = lastChangeToken.Split(';');

                    //var ticks = long.Parse(parts[3]);            // 638496224000000000
                    //var utcDate = new DateTime(ticks, DateTimeKind.Utc);

                    //_logger.LogInformation($"UTC Time: {utcDate:o}");

                    if (lastChangeToken != null && !String.IsNullOrEmpty(lastChangeToken))
                    {
                        changeQuery.ChangeTokenStart = new ChangeTokenOptions(lastChangeToken);
                    }


                    _logger.LogInformation($"SharePoint List Name: {targetList.Title}");

                    var changes = await targetList.GetChangesAsync(changeQuery);

                    if (changes.Any())
                    {
                        await _services.SaveLatestChangeTokenAsync(changes.Last().ChangeToken, resourceId, _settings.AzureBlobStorageConnectionString, _settings.BlobContainerList);
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
                                    var resultAsync = await _services.DossierMaitreCreation(clientContext, pnpCoreContext, changeItem.ItemId, targetList,_settings.deploymentEnv);
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
                        try { _services.Mailer(clientContext, pnpCoreContext, groupedItemsAsync, singleItemsAsync); }
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


