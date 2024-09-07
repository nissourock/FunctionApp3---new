using Azure.Storage.Blobs;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using Microsoft.Extensions.Logging;
using PnP.Core.Model;
using PnP.Core.Services;
using static FunctionApp3.Models;
using System.Globalization;
using System.Collections;



namespace FunctionApp3
{
    public class Services
    {
        public static async Task<object?> DossierMaitreCreation(ClientContext clientContext, PnPContext pnpCoreContext, int itemID, PnP.Core.Model.SharePoint.IList targetList, Microsoft.Extensions.Logging.ILogger _logger)
        {

            try
            {
                {
                    var item = await targetList.Items.GetByIdAsync(itemID, i => i.Title, i => i.All);
                    //var JSONtoDeserialize = myItem.Values["JSON"]?.ToString().Trim('"').Replace("'", "\"");
                    var JSONtoDeserialize = item.Values["JSON"]?.ToString().Trim('"').Replace("{\'", "{\"").Replace("\'}", "\"}").Replace("\':\'", "\":\"").Replace("\',\'", "\",\"").Replace("\n", "");
                    var ListName = item.Values["ListName"]?.ToString();
                    var pays = item.Values["FolderPathDestination"]?.ToString();
                    var statutWebhook = item.Values["Statut_x0020_webhook"]?.ToString();
                    var UserEmail = item.Values["creatorEmail"]?.ToString();
                    var creationMode = item.Values["ModeCreation"]?.ToString();

                    JSONMetadataFieldInList itemMetadata = JsonConvert.DeserializeObject<JSONMetadataFieldInList>(JSONtoDeserialize);

                    var documentLibrary2 = pnpCoreContext.Web.Lists.QueryProperties(l => l.Title, l => l.Fields).FirstOrDefault(list => list.Title == ListName);

                    await documentLibrary2.RootFolder.LoadAsync(f => f.ServerRelativeUrl);

                    // Manually convert semicolon-separated fields into arrays
                    //People Fields
                    string[]? emetteursList = itemMetadata.Emetteurs?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray();
                    string[]? verificateursList = itemMetadata.Verificateurs?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray(); ;
                    string[]? motsClesList = itemMetadata.MotsCles?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray(); ;
                    string[]? approbateurs = itemMetadata.Approbateurs?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray();
                    string[]? intervenantPAO = itemMetadata.IntervenantPAO?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray(); ;
                    string[]? intervenantTTX = itemMetadata.IntervenantTTX?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray();

                    //single people field
                    var chef = (string)null;
                    var secretaire = (string)null;
                    if (itemMetadata.Chef != "")
                    {
                        chef = itemMetadata.Chef?.Split('|')[2];
                    }
                    if (itemMetadata.Secretaire != "") 
                    { 
                        secretaire = itemMetadata.Secretaire?.Split('|')[2]; 
                    }

                    //Term store ID Fields
                    int[]? siteList = itemMetadata.Site?.Split(';').Where(v => int.TryParse(v, out _)).Select(int.Parse).ToArray();
                    int[]? motsCles = itemMetadata.MotsCles?.Split(';').Where(v => int.TryParse(v, out _)).Select(int.Parse).ToArray(); ;


                    //Lookup fields ID
                    int? codePays = int.TryParse(itemMetadata.CodePays, out var tempCodePays) ? tempCodePays : null;
                    int? societe = int.TryParse(itemMetadata.Societe, out var tempSociete) ? tempSociete : null;
                    int? codeProjet = int.TryParse(itemMetadata.CodeProjet, out var tempCodeProjet) ? tempCodeProjet : null;
                    int? client = int.TryParse(itemMetadata.Client, out var tempClient) ? tempClient : null;
                    int? departement = int.TryParse(itemMetadata.Departement, out var tempDepartement) ? tempDepartement : null;

                    //managed metadata
                    string specialiteMetier = itemMetadata.SpecialiteMetier;
                    string typeDeDocument = itemMetadata.TypeDeDocument;

                    //Datefield
                    string dateSouhaite = itemMetadata.DateSouhaite;
                    string isoDate = "";
                    var date = DateTime.UtcNow;
                    if (dateSouhaite != null)
                    {
                        try
                        {
                            date = DateTime.ParseExact(dateSouhaite, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                            isoDate = date.ToString("yyyy-MM-dd");
                        }
                        catch { isoDate = ""; }


                    }

                    //plain strings

                    string title = itemMetadata.Title;
                    string source = itemMetadata.Source;
                    string numeroDossier = itemMetadata.NumeroDossier;
                    string activite = itemMetadata.Activite;
                    string bibliotheque = itemMetadata.Bibliotheque;
                    string langue = itemMetadata.Langue;
                    string dossierProjet = itemMetadata.DossierProjet;
                    string etat = itemMetadata.Etat;
                    string phaseEtude = itemMetadata.PhaseEtude;
                    string revision = itemMetadata.Revision;
                    string objetRevision = itemMetadata.ObjetRevision;
                    string intituleProjet = itemMetadata.IntituleProjet;
                    string description = itemMetadata.Description;

                    // integer values
                    int? nombrePage = int.TryParse(itemMetadata.NombrePage, out var tempNombrePage) ? tempNombrePage : null;
                    int? nombreAnnexes = int.TryParse(itemMetadata.NombreAnnexes, out var tempNombreAnnexes) ? tempNombreAnnexes : null;


                    //lah yester 1
                    //commented part starts here
                    string documentSetTemplate2 = "TEMPLATE";

                    var camlQuery2 = new CamlQueryOptions
                    {
                        ViewXml = $@"<View>
                                  <ViewFields>
                                    <FieldRef Name='FileRef' />
                                    <FieldRef Name='FileLeafRef' />
                                  </ViewFields>
                                  <Query>
                                    <Where>
                                      <Eq>
                                        <FieldRef Name='FileLeafRef'/>
                                        <Value Type='Text'>{documentSetTemplate2}</Value>
                                      </Eq>
                                    </Where>
                                  </Query>
                                </View>",


                    };

                    // Load the items using the CAML query
                    await documentLibrary2.LoadItemsByCamlQueryAsync(camlQuery2);

                    // Access the loaded items
                    var documentSetItem2 = documentLibrary2.Items.AsRequested().FirstOrDefault();

                    if (documentSetItem2 != null)
                    {
                        // Get the server-relative URL of the Document Set folder
                        var sourceUrl2 = documentSetItem2["FileRef"].ToString();
                        var targetUrl2 = "";
                        // Log the source URL for debugging
                        _logger.LogInformation($"Source URL: {sourceUrl2}");
                        if (ListName == "Livrables Affaires") { targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{pays}{dossierProjet}/{title}"; }
                        if (ListName == "Hors livrables") { targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{title}"; }
                        if (ListName == "Anciens Livrables") { targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{pays}{title}"; }
                        if (ListName == "Production de documents") { targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{title}"; }
                        // Define the target URL
                        // Log the target URL for debugging
                        _logger.LogInformation($"Target URL: {targetUrl2}");

                        // Perform the copy operation for the folder (Document Set)
                        var sourceFolder2 = pnpCoreContext.Web.GetFolderByServerRelativeUrl(sourceUrl2);
                        try
                        {
                            sourceFolder2.CopyTo(targetUrl2, new PnP.Core.Model.SharePoint.MoveCopyOptions
                            {
                                KeepBoth = false,
                                RetainEditorAndModifiedOnMove = true,


                            });

                        }
                        catch (Exception e)
                        {
                            _logger.LogError(e.Message);
                        }



                        var copiedDocumentSetFolder = pnpCoreContext.Web.GetFolderByServerRelativeUrl(targetUrl2);
                        await copiedDocumentSetFolder.EnsurePropertiesAsync(f => f.ListItemAllFields, f => f.ParentFolder);
                        var copiedDocumentSetItemToRename = copiedDocumentSetFolder.ListItemAllFields;
                        var copiedDocumentSetListItem = await documentLibrary2.Items.GetByIdAsync(copiedDocumentSetItemToRename.Id, i => i.Title, i => i.All);

                        //   var copiedDocumentSetItem = documentLibrary.Items.AsRequested().FirstOrDefault(ds => ds["FileRef"].ToString().Equals(targetUrl, StringComparison.OrdinalIgnoreCase));
                        if (copiedDocumentSetFolder != null)
                        {
                            var CSOMlist = clientContext.Web.Lists.GetByTitle(ListName);
                            var CSOMlistItem = CSOMlist.GetItemById(copiedDocumentSetItemToRename.Id);
                            copiedDocumentSetItemToRename["FileLeafRef"] = title;
                            copiedDocumentSetItemToRename["Title"] = title;

                            if (ListName == "Livrables Affaires")
                            {
                                // string  numeroDossier  "Num_x00e9_ro_x0020_Dossier"
                                if (description != null)
                                {
                                    copiedDocumentSetListItem["DocumentSetDescription"] = description;
                                }

                            }

                            if (ListName == "Production de documents")
                            {

                                //  numeroDossier  "Num_x00e9_ro_x0020_Dossier" //form
                                if (activite != null) { copiedDocumentSetListItem["Activit_x00e9_"] = activite; }


                                //  bibliotheque "Biblioth_x00e8_que_x0020_Cible"
                                if (bibliotheque != null) { copiedDocumentSetListItem["Biblioth_x00e8_que_x0020_Cible"] = bibliotheque; }

                                //  langue  "Langue"
                                if (langue != null) { copiedDocumentSetListItem["Langue"] = langue; }

                                //  dateSouhaite "Date_x0020_Souhait_x00e9_e"
                                //     if (isoDate != "") { copiedDocumentSetListItem["Date_x0020_Souhait_x00e9_e"] = date; }
                                copiedDocumentSetListItem["Activit_x00e9_"] = activite;
                                //  nombrePage "Nombre_x0020_de_x0020_Pages" //form
                                if (nombrePage != null)
                                { copiedDocumentSetListItem["Nombre_x0020_de_x0020_Pages"] = nombrePage; }
                                if (nombreAnnexes != null)
                                { copiedDocumentSetListItem["Nombre_x0020_Annexes"] = nombreAnnexes; }
                                //  nombreAnnexes "Nombre_x0020_de_x0020_Pages"  //form
                                if (numeroDossier != null)

                                {
                                    var parentItemCount = copiedDocumentSetFolder.ParentFolder.ItemCount;
                                    string result = $"{DateTime.Now.Year}-{parentItemCount + 1}";
                                    copiedDocumentSetListItem["Num_x00e9_ro_x0020_Dossier"] = result;
                                }

                                if (etat != null) { copiedDocumentSetListItem["Etat"] = etat; }

                                if (phaseEtude != null)
                                {
                                    copiedDocumentSetListItem["Phase_x0020_d_x0027__x00e9_tude"] = phaseEtude;
                                }
                                //  phaseEtude "Phase_x0020_d_x0027__x00e9_tude"

                            }

                            if (ListName == "Hors livrables")
                            {
                                //  numeroDossier  "Num_x00e9_ro_x0020_Dossier"
                                if (activite != null)
                                {
                                    copiedDocumentSetListItem["Activit_x00e9_"] = activite;
                                }

                                //  bibliotheque "Biblioth_x00e8_que_x0020_Cible"
                                if (bibliotheque != null)
                                { copiedDocumentSetListItem["Biblioth_x00e8_que_x0020_Cible"] = bibliotheque; }
                                //  langue  "Langue"
                                if (langue != null) { copiedDocumentSetListItem["Langue"] = langue; }
                                //  dateSouhaite "Date_x0020_Souhait_x00e9_e"

                                //  nombrePage "Nombre_x0020_de_x0020_Pages"
                                //  nombreAnnexes "Nombre_x0020_de_x0020_Pages"

                            }


                            if (ListName == "Anciens Livrables")
                            {

                                //  numeroDossier  "Num_x00e9_ro_x0020_Dossier"

                                //  activite "Activit_x00e9_"
                                if (activite != null)
                                {
                                    copiedDocumentSetListItem["Activit_x00e9_"] = activite;
                                }

                            };

                            if (ListName != "Anciens Livrables") { copiedDocumentSetListItem["Statut"] = "Création"; } else { copiedDocumentSetListItem["Statut"] = "Archivé"; }
                            //new pays implementation for missing code pays but libellepays value exists
                            if(pays != "")
                            {

                                try
                                {
                                    var paysID = Services.GetItemIdsBySiteTitles(clientContext, "Codes Pays", pays.Replace("/", ""), "Libell_x00e9_");
                                    CSOMlistItem["CodePaysd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)Int32.Parse(paysID)};
                                    CSOMlistItem.SystemUpdate();
                                    clientContext.ExecuteQuery();
                                }
                                catch (Exception e) { _logger.LogError(e.Message); }



                                if (ListName != "Anciens Livrables" && ListName != "Hors livrables")
                                {
                                    try
                                    {
                                        var paysID = Services.GetItemIdsBySiteTitles(clientContext, "Codes Pays", pays.Replace("/", ""), "Libell_x00e9_");
                                        CSOMlistItem["CodePayscd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)Int32.Parse(paysID) };
                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                    }
                                    catch (Exception e) { _logger.LogError(e.Message); }

                                }

                            }

                            //Old implementation for pays when codepays is present
                            //if (codePays.HasValue)
                            {


                                try
                                {
                                    CSOMlistItem["CodePaysd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)codePays };
                                    CSOMlistItem.SystemUpdate();
                                    clientContext.ExecuteQuery();
                                }
                                catch (Exception e) { _logger.LogError(e.Message); }



                                if (ListName != "Anciens Livrables" && ListName != "Hors livrables")
                                {
                                    try
                                    {
                                        CSOMlistItem["CodePayscd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)codePays };
                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                    }
                                    catch (Exception e) { _logger.LogError(e.Message); }

                                }

                            }

                            if (societe.HasValue)
                            {
                                CSOMlistItem.RefreshLoad();
                                if (ListName != "Anciens Livrables")
                                {
                                    CSOMlistItem["Soci_x00e9_t_x00e9_0"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)societe };


                                }

                                else
                                {
                                    CSOMlistItem["Soci_x00e9_t_x00e9_"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)societe };

                                }
                                try
                                {

                                    CSOMlistItem.SystemUpdate();
                                    clientContext.ExecuteQuery();
                                }
                                catch (Exception e) { _logger.LogError(e.Message); }
                            }
                            if (departement.HasValue)
                            {

                                if (ListName == "Production de documents" | ListName == "Hors livrables")
                                {
                                    CSOMlistItem["D_x00e9_partement"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)departement };
                                    try
                                    {

                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                    }
                                    catch (Exception e) { _logger.LogError(e.Message); }

                                }
                            }

                            if (codeProjet.HasValue)
                            {
                                CSOMlistItem.RefreshLoad();
                                if (ListName == "Anciens Livrables" | ListName == "Hors livrables")
                                {
                                    CSOMlistItem["CodeProjetd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)codeProjet };


                                }
                                else
                                {

                                    CSOMlistItem["CodeProjet"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)codeProjet };



                                    CSOMlistItem["CodeProjetd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)codeProjet };

                                };
                                if (ListName == "Hors livrables")
                                {

                                    CSOMlistItem["CodeLibelleProjet"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)codeProjet };

                                }
                                try
                                {

                                    CSOMlistItem.SystemUpdate();
                                    clientContext.ExecuteQuery();
                                }
                                catch (Exception e) { _logger.LogError(e.Message); }

                            }
                            if (client.HasValue)
                            {
                                if (ListName == "Anciens Livrables" | ListName == "Hors livrables")
                                {
                                    try
                                    {
                                        CSOMlistItem["Clientd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)client };
                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                    }
                                    catch (Exception e) { _logger.LogError(e.Message); }

                                }
                                else if (ListName != "Hors livrables")
                                {
                                    try
                                    {
                                        CSOMlistItem["Clientd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)client };
                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                    }
                                    catch (Exception e) { _logger.LogError(e.Message); }
                                    try
                                    {
                                        CSOMlistItem["Client"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)client };
                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                    }
                                    catch (Exception e) { _logger.LogError(e.Message); }

                                };
                            }

                            if (siteList.Length > 0)
                            {
                                var lookupValues = new ArrayList();

                                //var lookupValue1 = new FieldLookupValue { LookupId = 5 };
                                //lookupValues.Add(lookupValue1);

                                foreach (int id in siteList)
                                {
                                    var lookupValue = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = id };
                                    lookupValues.Add(lookupValue);
                                    if (ListName == "Livrables Affaires" | ListName == "Hors livrables")
                                    {
                                        CSOMlistItem["Site"] = lookupValues.ToArray();

                                    }
                                    if (ListName == "Anciens Livrables")
                                    {
                                        CSOMlistItem["Site_x0020_GED"] = lookupValues.ToArray();

                                    }
                                }
                                try
                                {
                                    CSOMlistItem.SystemUpdate();
                                    clientContext.ExecuteQuery();
                                }
                                catch (Exception e) { _logger.LogError(e.Message); }
                            }
                            if (motsCles.Length > 0)
                            {
                                if (ListName == "Livrables Affaires")
                                {
                                    var lookupValues = new ArrayList();
                                    foreach (int id in motsCles)
                                    {
                                        var lookupValue = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = id };
                                        lookupValues.Add(lookupValue);


                                        // (copiedDocumentSetListItem["Mots_x0020_Cl_x00e9_s"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldLookupValue(id));
                                    }
                                    try
                                    {
                                        CSOMlistItem["Mots_x0020_Cl_x00e9_s"] = lookupValues.ToArray();

                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                    }
                                    catch (Exception e) { _logger.LogError(e.Message); }
                                }

                            }

                            //user fields
                            if (!string.IsNullOrEmpty(secretaire))
                            {
                                try
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));
                                    // ISharePointUser secretaireUser = pnpCoreContext.Web.SiteUsers.AsRequested().Where(i => i.Mail == secretaire).FirstOrDefault();
                                    var secretaireUser = await pnpCoreContext.Web.EnsureUserAsync(secretaire);
                                    //await secretaireUser.LoadAsync(p => p.Id, p => p.AadObjectId, p => p.All);


                                    copiedDocumentSetListItem["Secretaired"] = new PnP.Core.Model.SharePoint.FieldUserValue(secretaireUser);
                                    if (ListName != "Anciens Livrables")
                                    {

                                        copiedDocumentSetListItem["Secr_x00e9_taire"] = new PnP.Core.Model.SharePoint.FieldUserValue(secretaireUser);
                                    }


                                    copiedDocumentSetListItem.SystemUpdate();

                                }
                                catch (Exception e) { _logger.LogError(e.Message); }

                            }
                            if (!string.IsNullOrEmpty(chef))
                            {
                                try
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));
                                    // ISharePointUser secretaireUser = pnpCoreContext.Web.SiteUsers.AsRequested().Where(i => i.Mail == secretaire).FirstOrDefault();

                                    var chefUser = await pnpCoreContext.Web.EnsureUserAsync(chef);

                                    copiedDocumentSetListItem["Chef_x0020_de_x0020_Projet_x002F_Coordinateur_x0020_d"] = new PnP.Core.Model.SharePoint.FieldUserValue(chefUser);
                                    if (ListName != "Anciens Livrables")
                                    {

                                        copiedDocumentSetListItem["Chef_x0020_de_x0020_projet_x0020__x002F__x0020_Coordinateur"] = new PnP.Core.Model.SharePoint.FieldUserValue(chefUser);
                                    }



                                    copiedDocumentSetListItem.SystemUpdate();

                                }
                                catch (Exception e) { _logger.LogError(e.Message); }


                            }
                            try
                            {
                                foreach (string verificateur in verificateursList)
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));


                                    var verificateurUser = await pnpCoreContext.Web.EnsureUserAsync(verificateur);
                                    (copiedDocumentSetListItem["V_x00e9_rificateurs"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldUserValue(verificateurUser));
                                }


                                copiedDocumentSetListItem.SystemUpdate();

                            }
                            catch (Exception e) { _logger.LogError(e.Message); }
                            try
                            {
                                foreach (string approbateur in approbateurs)
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));


                                    var ApprobateursUser = await pnpCoreContext.Web.EnsureUserAsync(approbateur);
                                    (copiedDocumentSetListItem["Approbateurs"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldUserValue(ApprobateursUser));
                                }


                                copiedDocumentSetListItem.SystemUpdate();

                            }
                            catch (Exception e) { _logger.LogError(e.Message); }
                            try
                            {
                                foreach (string emetteur in emetteursList)
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));


                                    var emetteurUser = await pnpCoreContext.Web.EnsureUserAsync(emetteur);
                                    (copiedDocumentSetListItem["Emetteur"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldUserValue(emetteurUser));
                                }


                                copiedDocumentSetListItem.SystemUpdate();

                            }
                            catch (Exception e) { _logger.LogError(e.Message); }



                            try
                            {
                                foreach (string verificateur in intervenantPAO)
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));


                                    var verificateurUser = await pnpCoreContext.Web.EnsureUserAsync(verificateur);
                                    (copiedDocumentSetListItem["V_x00e9_rificateurs"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldUserValue(verificateurUser));
                                }


                                copiedDocumentSetListItem.SystemUpdate();

                            }
                            catch (Exception e) { _logger.LogError(e.Message); }
                            try
                            {
                                foreach (string verificateur in intervenantPAO)
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));


                                    var verificateurUser = await pnpCoreContext.Web.EnsureUserAsync(verificateur);
                                    (copiedDocumentSetListItem["V_x00e9_rificateurs"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldUserValue(verificateurUser));
                                }


                                copiedDocumentSetListItem.SystemUpdate();

                            }
                            catch (Exception e) { _logger.LogError(e.Message); }
                            //lookup fields

                            await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));
                            // ISharePointUser secretaireUser = pnpCoreContext.Web.SiteUsers.AsRequested().Where(i => i.Mail == secretaire).FirstOrDefault();
                            if (UserEmail != null && UserEmail != "")
                            {
                                var creatorEmail = await pnpCoreContext.Web.EnsureUserAsync(UserEmail);
                                copiedDocumentSetListItem["Author"] = new PnP.Core.Model.SharePoint.FieldUserValue(creatorEmail);
                                copiedDocumentSetListItem["Editor"] = new PnP.Core.Model.SharePoint.FieldUserValue(creatorEmail);

                            }
                            try
                            {

                                await copiedDocumentSetListItem.SystemUpdateAsync();
                                await copiedDocumentSetItemToRename.UpdateAsync();
                            }
                            catch (Exception e) { _logger.LogError(e.Message); }
                            
                            //Only non system update update. To increment version number and update the modified timestamp
                            copiedDocumentSetListItem.Update();
                            copiedDocumentSetItemToRename.Update();

                            

                            //managed metadata fields

                            var timeZone = TimeZoneInfo.FindSystemTimeZoneById("Romance Standard Time");
                            if (dateSouhaite != "")
                            {
                                _logger.LogInformation(DateTime.Parse(dateSouhaite, new CultureInfo("fr-FR")).ToString(new CultureInfo("fr-FR")));
                                var formValues3 = new List<ListItemFormUpdateValue>
                                                                                     {
                                                                            new ListItemFormUpdateValue() { FieldName = "Date_x0020_Souhait_x00e9_e" , FieldValue = DateTime.Parse(dateSouhaite, new CultureInfo("fr-FR")).ToString(new CultureInfo("fr-FR"))}

                                                                                     };
                                try
                                {
                                    CSOMlistItem.ValidateUpdateListItem(formValues3, false, "", false, false, "");
                                    clientContext.ExecuteQuery();
                                }
                                catch (Exception e) { _logger.LogError(e.Message); }

                            }
                            //update creation and modified date and time 
                            //_logger.LogInformation(TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, timeZone).ToString(new CultureInfo("fr-FR")));
                            var formValues5 = new List<ListItemFormUpdateValue>
                            {
                            new ListItemFormUpdateValue() {FieldName= "Created", FieldValue = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, timeZone).ToString(new CultureInfo("fr-FR")) }
                            } ;
                            try
                            {
                                CSOMlistItem.ValidateUpdateListItem(formValues5, false, "", false, false, "");
                                clientContext.ExecuteQuery();

                            }
                            catch (Exception e) { _logger.LogError(e.Message); }


                            //Force update faulty columns in hors livrables 
                            if (ListName == "Hors livrables")
                            {
                                var formValues4 = new List<ListItemFormUpdateValue>
                                                                                     {
                                                                            new ListItemFormUpdateValue() { FieldName = "CodePaysd" , FieldValue = codePays.ToString()},
                                                                            new ListItemFormUpdateValue() { FieldName = "Clientd" , FieldValue = client.ToString()},


                                                                                     }
                                ;
                                //var lookupValues = new ArrayList();

                                //CSOMlistItem["Site"] = 
                                try
                                {
                                    CSOMlistItem.ValidateUpdateListItem(formValues4, false, "", false, false, "");
                                    clientContext.ExecuteQuery();

                                }
                                catch (Exception e) { _logger.LogError(e.Message); }


                            }

                            if (!string.IsNullOrEmpty(specialiteMetier) && ListName == "Production de documents")
                            {

                                var formValues1 = new List<ListItemFormUpdateValue>
                                                                                    {
                                                                            new ListItemFormUpdateValue() { FieldName = "Sp_x00e9_cialit_x00e9__x002f_m_x00e9_tier" , FieldValue = specialiteMetier}
                                                                                    };

                                try
                                {

                                    CSOMlistItem.ValidateUpdateListItem(formValues1, false, "", false, false, "");
                                    clientContext.ExecuteQuery();
                                }
                                catch (Exception e) { _logger.LogError(e.Message); }

                            }
                            else if (!string.IsNullOrEmpty(specialiteMetier))
                            {
                                var formValues1 = new List<ListItemFormUpdateValue>
                                                                                    {
                                                                            new ListItemFormUpdateValue() { FieldName = "Sp_x00e9_cialit_x00e9__x002F_m_x00e9_tier" , FieldValue = specialiteMetier}
                                                                                    };
                                try
                                {

                                    CSOMlistItem.ValidateUpdateListItem(formValues1, false, "", false, false, "");
                                    clientContext.ExecuteQuery();
                                }
                                catch (Exception e) { _logger.LogError(e.Message); }

                            }

                            if (!string.IsNullOrEmpty(typeDeDocument))
                            {

                                var formValues2 = new List<ListItemFormUpdateValue>
                                                                                     {
                                                                            new ListItemFormUpdateValue() { FieldName = "Typededocument" , FieldValue = typeDeDocument}
                                                                                     };
                                try
                                {
                                    CSOMlistItem.ValidateUpdateListItem(formValues2, false, "", false, false, "");
                                    clientContext.ExecuteQuery();

                                }
                                catch (Exception e) { _logger.LogError(e.Message); }



                            }
                            //numero dossier
                            
                            _logger.LogInformation($"Document Set copied and renamed to: {title}. Metadata successfully updated");
                            item["Statut_x0020_webhook"] = "Crée";
                            await item.UpdateAsync();
                            _logger.LogInformation($"Item status updated to created in og list");
                            return new {targetURL=targetUrl2, titleDossier = title, modecreation = creationMode};

                        }
                        else
                        {
                            _logger.LogError($"Failed to find the copied Document Set at {targetUrl2}.");
                            return null;
                        }
                    }
                    else
                    {
                        _logger.LogError($"Document Set named 'Tester' not found in 'Test webhook' document library.");
                        return null;
                    }


                    //                //commented part ends here

                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        public static async Task SaveLatestChangeTokenAsync(IChangeToken changeToken, string ressource, string blobconnection, string container)
        {
            // Get a reference to the Azure Storage Container
            //string blobconnection = "AzureWebJobsStorage";
            try
            {


                BlobContainerClient containerClient = new BlobContainerClient(blobconnection, container);
                // Get a reference to the Azure Storage Blob
                var blobClient = containerClient.GetBlobClient(ressource + ".txt");


                // Prepare the JSON content
                using (var mem = new MemoryStream())
                {

                    using (var sw = new StreamWriter(mem))
                    {

                        sw.WriteLine(changeToken.StringValue);
                        await sw.FlushAsync();

                        mem.Position = 0;

                        // Upload it into the target blob
                        await blobClient.UploadAsync(mem, overwrite: true);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");

            }

        }
        public static string FormatEmails(string emailList, string prefix)
        {
            if (!string.IsNullOrEmpty(emailList))
            {
                // Split by comma, trim spaces, and add the prefix to each element
                var processedList = emailList
                    .Split(',')
                    .Select(email => prefix + email.Trim())
                    .ToArray();

                // Concatenate them separated by ;
                return string.Join(";", processedList);
            }
            else return string.Empty;
        }
        public static string GetItemIdsBySiteTitles(ClientContext context, string listTitle, string values, string columnName)
        {
            if (string.IsNullOrEmpty(values))
                return string.Empty;
            try
            {

                var valuesArray = values.Split(',');
                var concatenatedIds = new List<string>();

                var list = context.Web.Lists.GetByTitle(listTitle);

                foreach (var value in valuesArray)
                {
                    // Load all items from the list


                    var camlQuery = new CamlQuery
                    {
                        ViewXml = $@"<View><ViewFields>
                                   <FieldRef Name='{columnName}'/>
                                   <FieldRef Name='Title'/>
                                   
                                 </ViewFields>
                                 <Query>
                                   <Where>
                                     <Contains>
                                       <FieldRef Name='{columnName}'/>
                                       <Value Type='text'>{value.Trim()}</Value>
                                     </Contains>
                                   </Where>
                                 </Query>
                                </View>",

                    };
                    
                    var items = list.GetItems(camlQuery);
                    context.Load(items);
                    context.ExecuteQuery();
                    var listItem = items.AsEnumerable().FirstOrDefault();
                    // Filter items based on "Site" column

                       if (listItem != null)
                    {
                        concatenatedIds.Add(listItem.Id.ToString());
                    }
                    else
                    {
                        Console.WriteLine($"List item with title '{value}' not found.");
                        concatenatedIds.Add("");
                    }
                }

                return string.Join(";", concatenatedIds);
            }
            catch { return string.Empty; }
        }
        public static async Task<string> GetLatestChangeTokenAsync(string ressource, string blobconnection, string container)
        {
            {
                BlobContainerClient containerClient = new BlobContainerClient(blobconnection, container);

                // Ensure the container exists
                await containerClient.CreateIfNotExistsAsync();

                // Verify if the container exists
                bool containerExists = await containerClient.ExistsAsync();
                if (!containerExists)
                {
                    throw new Exception($"Blob container '{container}' does not exist and could not be created.");
                }

                // Get a reference to the blob
                BlobClient blobClient = containerClient.GetBlobClient(ressource + ".txt");

                // Check if the blob exists
                bool blobExists = await blobClient.ExistsAsync();
                if (!blobExists)
                {
                    // Create the blob with initial content if it does not exist
                    string initialContent = "";
                    using (var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(initialContent)))
                    {
                        await blobClient.UploadAsync(ms);
                    }
                    return initialContent;
                }

                // Download the blob content
                var blobContent = await blobClient.DownloadContentAsync();
                var blobContentString = blobContent.Value.Content.ToString();
                return blobContentString;
            }
        }
    }
}
