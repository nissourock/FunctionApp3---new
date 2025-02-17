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
using Microsoft.SharePoint.Client.Utilities;
using FieldUserValue = Microsoft.SharePoint.Client.FieldUserValue;
using Microsoft.AspNetCore.JsonPatch.Internal;




namespace FunctionApp3
{
    public class Services
    {
        public static string ReplaceLastPart(string input, string replacement)
        {
            // Remove trailing slash if it exists
            string trimmedInput = input.TrimEnd('/');

            int lastSlashIndex = trimmedInput.LastIndexOf('/');
            if (lastSlashIndex == -1)
            {
                // No slashes found, return the replacement as the full string
                return replacement;
            }

            // Replace the part after the last meaningful slash
            return trimmedInput.Substring(0, lastSlashIndex + 1) + replacement;
        }
        public static async Task<dynamic?> DossierMaitreCreation(ClientContext clientContext, PnPContext pnpCoreContext, int itemID, PnP.Core.Model.SharePoint.IList targetList, Microsoft.Extensions.Logging.ILogger _logger)
        {

            try
            {
                {
                    var item = await targetList.Items.GetByIdAsync(itemID, i => i.Title, i => i.All);
                    // The next line transforms single quotes into double quotes for JSON parse, also fixes some potential power apps anomalies
                    var JSONtoDeserialize = item.Values["JSON"]?.ToString().Trim('"').Replace("        ", "").Replace("   ,", "").Replace("{\'", "{\"").Replace("\'}", "\"}").Replace("\':\'", "\":\"").Replace("\',\'", "\",\"").Replace("\n", "").Replace("   \'","\'");
                    var ListName = item.Values["ListName"]?.ToString();
                    var pays = item.Values["FolderPathDestination"]?.ToString();
                    //ppr 
                    var statutWebhook = item.Values["Statut_x0020_webhook"]?.ToString(); 
                    //prod
                    //var statutWebhook = item.Values["Statutwebhook"]?.ToString();
                    
                    var UserEmail = item.Values["creatorEmail"]?.ToString();
                    var creationMode = item.Values["ModeCreation"]?.ToString();
                    var operationIdentifier = item.Values["OperationIdentifier"]?.ToString(); 
                    bool isModified = false;



                    //item["Statutwebhook"] = "Création en cours";
                    item["Statut_x0020_webhook"] = "Création en cours";


                    await item.UpdateAsync();

                    _logger.LogInformation($"Updated status in og list");

                    JSONMetadataFieldInList itemMetadata = JsonConvert.DeserializeObject<JSONMetadataFieldInList>(JSONtoDeserialize);

                    _logger.LogInformation($"Successfully parsed JSON data");

                    var documentLibrary2 = pnpCoreContext.Web.Lists.QueryProperties(l => l.Title, l => l.Fields).FirstOrDefault(list => list.Title == ListName);

                    await documentLibrary2.RootFolder.LoadAsync(f => f.ServerRelativeUrl);

                    // Manually convert semicolon-separated fields into arrays
                    //People Fields
                    string?[]? emetteursList = itemMetadata.Emetteurs?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray();
                    string?[]? verificateursList = itemMetadata.Verificateurs?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray(); ;
                    string?[]? motsClesList = itemMetadata.MotsCles?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray(); ;
                    string?[]? approbateurs = itemMetadata.Approbateurs?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray();
                    string?[]? intervenantPAO = itemMetadata.IntervenantPAO?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray(); ;
                    string?[]? intervenantTTX = itemMetadata.IntervenantTTX?.Split(';').Select(e => e.Split('|').ElementAtOrDefault(2))
                            .Where(e => !string.IsNullOrEmpty(e))
                            .ToArray();

                    //single people field
                    var chef = (string?)null;
                    var secretaire = (string?)null;
                    if (itemMetadata.Chef != "")
                    {
                        chef = itemMetadata.Chef?.Split('|').Length>1 ? chef = itemMetadata.Chef?.Split('|')[2] : itemMetadata.Chef ;
                    }
                    if (itemMetadata.Secretaire != "") 
                    { 
                        secretaire = itemMetadata.Secretaire?.Split('|').Length > 1 ? secretaire = itemMetadata.Secretaire?.Split('|')[2] : itemMetadata.Secretaire; 
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
                            _logger.LogInformation($"Parsed date souhaitee date");
                        }
                        catch { isoDate = ""; }


                    }


                    //plain strings

                    string title = itemMetadata.Title.Replace("\\","").Trim();
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
                    string FilRefDoc = itemMetadata.FilRefDoc;
                    // integer values
                    int? nombrePage = int.TryParse(itemMetadata.NombrePage, out var tempNombrePage) ? tempNombrePage : null;
                    int? nombreAnnexes = int.TryParse(itemMetadata.NombreAnnexes, out var tempNombreAnnexes) ? tempNombreAnnexes : null;

                    _logger.LogInformation($"List: {ListName}, Title: {title}, parsed data correctly, initiating creation");

                    //lah yester 1
                    //commented part starts here
                    string documentSetTemplate2 = "TEMPLATE";

                    var camlQuery2 = new CamlQueryOptions
                    {
                        ViewXml = $@"<View>
                                  <ViewFields>
                                    <FieldRef Name='FileRef'/>
                                    <FieldRef Name='FileLeafRef'/>
                                  </ViewFields>
                                  <Query>
                                    <Where>
                                      <Eq>
                                        <FieldRef Name='FileLeafRef'/>
                                        <Value Type='Text'>{documentSetTemplate2}</Value>
                                      </Eq>
                                    </Where>
                                  </Query>
                                </View>"
                    };

                    // Load the items using the CAML query
                    await documentLibrary2.LoadItemsByCamlQueryAsync(camlQuery2);

                    // Access the loaded items
                    var documentSetItem2 = documentLibrary2.Items.AsRequested().FirstOrDefault();

                    if (documentSetItem2 != null)

                    {
                        _logger.LogInformation($"TEMPLATE Found");
                        // Get the server-relative URL of the Document Set folder
                        var sourceUrl2 = documentSetItem2["FileRef"].ToString();
                        var targetUrl2 = "";
                        // Log the source URL for debugging
                        _logger.LogInformation($"Source URL: {sourceUrl2}");
                        if (ListName == "Livrables Affaires")
                        {
                            var ProgID = "";

                            if (FilRefDoc != null && FilRefDoc != "")
                            {
                                try {
                                    var DocumentSetFolder = pnpCoreContext.Web.GetFolderByServerRelativeUrl($"{documentLibrary2.RootFolder.ServerRelativeUrl}/{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}");
                                    ProgID = DocumentSetFolder?.ProgID;
                                }
                                catch (Exception e)
                                { _logger.LogError(e.Message); }
                                
                                if (ProgID == "Sharepoint.DocumentSet") { 
                                    targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}";
                                    isModified = true;
                                }
                                else
                                try
                                {
                                    var test = FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1);
                                    targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}/{title}";
                                        
                                    }
                                catch (Exception e) { targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{pays}{dossierProjet}/{title}"; }

                                
                            }
                            else { targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{pays}{dossierProjet}/{title}"; }




                        }




                        if (ListName == "Production de documents") {
                            var ProgID = "";

                            if (FilRefDoc != null && FilRefDoc != "")
                            {
                                try
                                {
                                    var DocumentSetFolder = pnpCoreContext.Web.GetFolderByServerRelativeUrl($"{documentLibrary2.RootFolder.ServerRelativeUrl}/{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}");
                                    ProgID = DocumentSetFolder?.ProgID;
                                }
                                catch (Exception e)
                                { _logger.LogError(e.Message); }
                                if (ProgID == "Sharepoint.DocumentSet")
                                {
                                    targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}";
                                    isModified = true;
                                }
                                else
                                
                                    try
                                {
                                    targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}/{title}";
                                    var targetUrl3 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{dossierProjet}/{title}";
                                }
                                catch (Exception e) {  targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{dossierProjet}/{title}"; }


                            }
                            else { targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{dossierProjet}/{title}"; }


                        }


                        if (ListName == "Hors livrables")
                        {
                            var ProgID = "";
                            if (FilRefDoc != null && FilRefDoc != "")

                            {
                                try
                                {
                                    var tesurl2 = "";
                                    var testURL = $"{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}";
                                    var DocumentSetFolder = pnpCoreContext.Web.GetFolderByServerRelativeUrl(FilRefDoc);
                                    ProgID = DocumentSetFolder?.ProgID;
                                }
                                catch (Exception e)
                                { _logger.LogError(e.Message); }
                                if (ProgID == "Sharepoint.DocumentSet")
                                {
                                    targetUrl2 = FilRefDoc;
                                    isModified = true;
                                }
                                else targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{title}";

                            }
                            else
                                targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{title}";

                        }

                        if (ListName == "Anciens Livrables") {
                            var ProgID = "";

                            try
                            {
                                var DocumentSetFolder = pnpCoreContext.Web.GetFolderByServerRelativeUrl($"{documentLibrary2.RootFolder.ServerRelativeUrl}/{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}");
                                ProgID = DocumentSetFolder?.ProgID;
                            }
                            catch (Exception e)
                            { _logger.LogError(e.ToString()); }
                            if (ProgID == "Sharepoint.DocumentSet")
                            {
                                targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{FilRefDoc.Substring(FilRefDoc.IndexOf('/') + 1)}";
                                isModified = true;
                            }
                            else
                            { targetUrl2 = $"{documentLibrary2.RootFolder.ServerRelativeUrl}/{pays}{title}"; }
                            }
                        
                        // Define the target URL
                        // Log the target URL for debugging
                        _logger.LogInformation($"Target URL: {targetUrl2}");
                        
                        // Perform the copy operation for the folder (Document Set)
                        //breakpoint was here
                        var sourceFolder2 = pnpCoreContext.Web.GetFolderByServerRelativeUrl(sourceUrl2);
                        ;
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
                            // An error of type file already exists means the user intended modification, so we delete the original file and create a new one with the initial props and modified ones 
                            //(when the user modifies an element, a new entry is created in the dossier maitres création list and it contains all props )

                            //var errorObject = e.GetType().GetProperty("Error")?.GetValue(e);
                            //var errorMessage = errorObject?.GetType().GetProperty("Message")?.GetValue(errorObject) as string;
                            //if(errorMessage == "The destination file already exists.") {
                            //    var documentSetFoldertoDelete = pnpCoreContext.Web.GetFolderByServerRelativeUrl(targetUrl2);
                            //    documentSetFoldertoDelete.Delete();
                            //    clientContext.ExecuteQuery();
                            //    // the file is flagged as modified so as to send the correct email later on 
                            //    isModified = true;
                            //    sourceFolder2.CopyTo(targetUrl2, new PnP.Core.Model.SharePoint.MoveCopyOptions
                            //    {
                            //        KeepBoth = false,
                            //        RetainEditorAndModifiedOnMove = true,
                            //    });
                            //}

                            _logger.LogError(e.ToString());
                            
                            isModified = true;

                            _logger.LogWarning("");
                        }




                        var copiedDocumentSetFolder = pnpCoreContext.Web.GetFolderByServerRelativeUrl(targetUrl2);
                        await copiedDocumentSetFolder.EnsurePropertiesAsync(f => f.ListItemAllFields, f => f.ParentFolder);
                        var copiedDocumentSetItemToRename = copiedDocumentSetFolder.ListItemAllFields;
                        var copiedDocumentSetListItem = await documentLibrary2.Items.GetByIdAsync(copiedDocumentSetItemToRename.Id, i => i.Title, i => i.All);

                        //   var copiedDocumentSetItem = documentLibrary.Items.AsRequested().FirstOrDefault(ds => ds["FileRef"].ToString().Equals(targetUrl, StringComparison.OrdinalIgnoreCase));
                        if (copiedDocumentSetFolder != null)
                        {
                            if(isModified == true) { _logger.LogWarning("Ignore file copy error, initiating edit"); }
                            var CSOMlist = clientContext.Web.Lists.GetByTitle(ListName);
                            
                            var CSOMlistItem = CSOMlist.GetItemById(copiedDocumentSetListItem.Id);

                           if(isModified == true) {CSOMlistItem["FileLeafRef"] = title; };
                            
                            CSOMlistItem["Title"] = title;

                            try
                            {
                                CSOMlistItem.Update();
                                clientContext.ExecuteQuery();
                                copiedDocumentSetListItem.Load();
                                _logger.LogInformation($"Updated Title");
                            }
                            catch (Exception e) { _logger.LogError(e.ToString()); }

                            

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
                                if (numeroDossier == ""  )

                                {
                                    var parentItemCount = copiedDocumentSetFolder.ParentFolder.ItemCount;
                                    string result = $"{DateTime.Now.Year}-{parentItemCount + 1}";
                                    copiedDocumentSetListItem["Num_x00e9_ro_x0020_Dossier"] = result;
                                } else { copiedDocumentSetListItem["Num_x00e9_ro_x0020_Dossier"] = numeroDossier; }

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

                                
                                if (bibliotheque != null)
                                { copiedDocumentSetListItem["Biblioth_x00e8_que_x0020_Cible"] = bibliotheque; }
                              
                                if (langue != null) { copiedDocumentSetListItem["Langue"] = langue; }
                                

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

                            CSOMlistItem.RefreshLoad();
                            clientContext.ExecuteQuery();

                            if (ListName != "Anciens Livrables") { 
                                
                                CSOMlistItem["Statut"] = "Création";
                            } else {
                                
                                CSOMlistItem["Statut"] = "Archivé";
                            }
                            try
                            {
                                
                                CSOMlistItem.SystemUpdate();
                                clientContext.ExecuteQuery();
                            }
                            catch (Exception e) { _logger.LogError(e.Message); }

                            CSOMlistItem.RefreshLoad();
                            clientContext.ExecuteQuery();
                            //new pays implementation for missing code pays but libellepays value exists
                            if (pays != "" && pays !="/")
                            {
                                if (ListName != "Hors livrables")
                                {
                                    try
                                    {
                                        var paysValue = pays.Replace("/", "");
                                        //We use this helper method to get the "code pays" ID by providing the pays intitulé
                                        var paysID = Services.GetItemIdsBySiteTitles(clientContext, "Codes Pays", paysValue, "Libell_x00e9_", _logger);
                                        CSOMlistItem["CodePaysd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)Int32.Parse(paysID) };
                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                        _logger.LogInformation($"Code pays OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }

                                }

                                if (ListName != "Anciens Livrables" && ListName != "Hors livrables")
                                {
                                    try
                                    {
                                        var paysID = Services.GetItemIdsBySiteTitles(clientContext, "Codes Pays", pays.Replace("/", ""), "Libell_x00e9_",_logger);
                                        CSOMlistItem["CodePayscd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)Int32.Parse(paysID) };
                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                        _logger.LogInformation($"Codepayscd OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }

                                }
                                //Force update faulty columns in hors livrables 
                                if (ListName == "Hors livrables")
                                {
                                    var formValues4 = new List<ListItemFormUpdateValue>
                                                                                     {
                                                                            new ListItemFormUpdateValue() { FieldName = "CodePaysd" , FieldValue = codePays.ToString()},

                                                                                     };

                                    try
                                    {
                                        CSOMlistItem.ValidateUpdateListItem(formValues4, false, "", false, false, "");
                                        clientContext.ExecuteQuery();

                                        _logger.LogInformation($"Codepaysd OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }

                                }
                            } else
                            //This is to set the codepays on hors livrables and production de document 
                            {
                                if (codeProjet.HasValue) 
                                {
                                    try
                                    {
                                        if (ListName == "Hors livrables"| ListName == "Production de documents")
                                        {
                                            //We use this helper method to get the "code pays" ID by providing the pays intitulé
                                            var paysFromProjet = Services.GetPaysIDFromProjet(clientContext, "Codes Projets", codeProjet, _logger);
                                            
                                            //We use this helper method to get the "code pays" ID by providing the pays intitulé
                                            var paysID = Services.GetItemIdsBySiteTitles(clientContext, "Codes Pays", paysFromProjet, "Libell_x00e9_", _logger);

                                            var formValuesPaysHorslivrables = new List<ListItemFormUpdateValue>
                                                                                     {
                                                                            new ListItemFormUpdateValue() { FieldName = "CodePaysd" , FieldValue = paysID.ToString()},

                                                                                     };
                                            CSOMlistItem.ValidateUpdateListItem(formValuesPaysHorslivrables, false, "", false, false, "");
                                            clientContext.ExecuteQuery();
                                            _logger.LogInformation($"Code Projet OK");
                                        }
                                            
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }
                                };
                            }

                            //Old implementation for pays when codepays is present
                            //if (codePays.HasValue)
                            //{
                            //    CSOMlistItem.RefreshLoad();

                            //    try
                            //    {
                            //        CSOMlistItem["CodePaysd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)codePays };
                            //        CSOMlistItem.SystemUpdate();
                            //        clientContext.ExecuteQuery();
                            //    }
                            //    catch (Exception e) { _logger.LogError(e.Message); }



                            //    if (ListName != "Anciens Livrables" && ListName != "Hors livrables")
                            //    {
                            //        try
                            //        {
                            //            CSOMlistItem["CodePayscd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)codePays };
                            //            CSOMlistItem.SystemUpdate();
                            //            clientContext.ExecuteQuery();
                            //        }
                            //        catch (Exception e) { _logger.LogError(e.Message); }

                            //    }

                            //}

                            if (societe.HasValue)
                            {
                               CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
                                if (ListName != "Anciens Livrables")
                                {
                                    CSOMlistItem["Soci_x00e9_t_x00e9_0"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)societe };
                                    try
                                    {

                                        CSOMlistItem.SystemUpdate();
                                        clientContext.ExecuteQuery();
                                        _logger.LogInformation($"Societe OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }

                                }

                                else
                                {
                                    //(copiedDocumentSetListItem["Soci_x00e9_t_x00e9_"] as IFieldLookupValue).LookupId = (int)societe;
                                    //copiedDocumentSetListItem.Update();
                                    //CSOMlistItem["Soci_x00e9_t_x00e9_"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)societe };
                                    var SocieteFormValues = new List<ListItemFormUpdateValue>
                                                                                     {
                                                                            new ListItemFormUpdateValue() { FieldName = "Soci_x00e9_t_x00e9_" ,FieldValue= societe.ToString()+"#UGS"}

                                                                                     };
                                    try
                                    {
                                        CSOMlistItem.ValidateUpdateListItem(SocieteFormValues, true, "", false, false, "");
                                      //CSOMlistItem.Update();
                                      clientContext.ExecuteQuery();
                                        _logger.LogInformation($"Societe OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }

                                }
                                
                                
                            }
                            if (departement.HasValue)
                            {
                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
                                if (ListName == "Production de documents" | ListName == "Hors livrables")
                                {
                                    CSOMlistItem["D_x00e9_partement"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)departement };
                                    try
                                    {

                                        CSOMlistItem.Update();
                                        clientContext.ExecuteQuery();
                                        _logger.LogInformation($"Departement OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }

                                }
                            }

                            if (codeProjet.HasValue)
                            {
                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
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
                                    _logger.LogInformation($"Code projet OK");
                                }
                                catch (Exception e) { _logger.LogError(e.ToString()); }

                            }
                            if (client.HasValue)
                            {

                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
                                if (ListName == "Anciens Livrables" | ListName == "Hors livrables")
                                {
                                    try
                                    {
                                        CSOMlistItem["Clientd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)client };
                                        CSOMlistItem.Update();
                                        clientContext.ExecuteQuery();
                                        _logger.LogInformation($"Clientd OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }

                                }
                                else if (ListName != "Hors livrables" && ListName != "Anciens Livrables")
                                {
                                    try
                                    {
                                        CSOMlistItem["Clientd"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)client };
                                        CSOMlistItem["Client"] = new Microsoft.SharePoint.Client.FieldLookupValue { LookupId = (int)client };
                                        CSOMlistItem.Update();
                                        clientContext.ExecuteQuery();
                                        _logger.LogInformation($"Clientd and client OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }
                                    

                                };
                            }

                            if (siteList.Length > 0)
                            {
                                CSOMlistItem.RefreshLoad();

                                
                                clientContext.ExecuteQuery();
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
                                    _logger.LogInformation($"Sites OK");
                                }
                                catch (Exception e) { _logger.LogError(e.ToString()); }
                            }
                            if (motsCles.Length > 0)
                            {

                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
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
                                        _logger.LogInformation($"Mots cles OK");
                                    }
                                    catch (Exception e) { _logger.LogError(e.ToString()); }
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
                                    _logger.LogInformation($"Secretaire OK");
                                }
                                catch (Exception e) { _logger.LogError(e.ToString()); }

                            }
                            if (!string.IsNullOrEmpty(chef))
                            {
                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
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
                                    _logger.LogInformation($"Chef OK");

                                }
                                catch (Exception e) { _logger.LogError(e.ToString()); }


                            }
                            try
                            {
                                foreach (string verificateur in verificateursList)
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));


                                    var verificateurUser = await pnpCoreContext.Web.EnsureUserAsync(verificateur);
                                    (copiedDocumentSetListItem["V_x00e9_rificateurs"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldUserValue(verificateurUser));
                                }


                                if (verificateursList.Length > 0) { copiedDocumentSetListItem.SystemUpdate();
                                    _logger.LogInformation($"Verificateurs OK");
                                }

                            }
                            catch (Exception e) { _logger.LogError(e.ToString()); }
                            try
                            {
                                foreach (string approbateur in approbateurs)
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));


                                    var ApprobateursUser = await pnpCoreContext.Web.EnsureUserAsync(approbateur);
                                    (copiedDocumentSetListItem["Approbateurs"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldUserValue(ApprobateursUser));
                                }

                                if (approbateurs.Length > 0) { copiedDocumentSetListItem.SystemUpdate();
                                    _logger.LogInformation($"Approbateurs OK");
                                }
                                

                            }
                            catch (Exception e) { _logger.LogError(e.ToString()); }
                            try
                            {
                                foreach (string emetteur in emetteursList)
                                {
                                    await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));


                                    var emetteurUser = await pnpCoreContext.Web.EnsureUserAsync(emetteur);
                                    (copiedDocumentSetListItem["Emetteur"] as IFieldValueCollection).Values.Add(new PnP.Core.Model.SharePoint.FieldUserValue(emetteurUser));
                                }

                                if (emetteursList.Length> 0) { copiedDocumentSetListItem.SystemUpdate();
                                    _logger.LogInformation($"Emetteurs OK");
                                }
                                

                            }
                            catch (Exception e) { _logger.LogError(e.ToString()); }


                            //lookup fields
                            try
                            {
                                await pnpCoreContext.Web.LoadAsync(p => p.SiteUsers.QueryProperties(i => i.Mail));
                                //copiedDocumentSetListItem.SystemUpdate();
                                //await copiedDocumentSetItemToRename.SystemUpdateAsync();
                            }
                            catch (Exception e) { _logger.LogError(e.ToString()); }


                            // ISharePointUser secretaireUser = pnpCoreContext.Web.SiteUsers.AsRequested().Where(i => i.Mail == secretaire).FirstOrDefault();
                            if (UserEmail != null && UserEmail != "")
                            {
                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
                                var creatorEmail = await pnpCoreContext.Web.EnsureUserAsync(UserEmail);
                                //copiedDocumentSetListItem["Author"] = new PnP.Core.Model.SharePoint.FieldUserValue(creatorEmail);
                                //copiedDocumentSetListItem["Editor"] = new PnP.Core.Model.SharePoint.FieldUserValue(creatorEmail);

                                CSOMlistItem["Author"] = new FieldUserValue() { LookupId= creatorEmail.Id};
                                CSOMlistItem["Editor"] = new FieldUserValue() { LookupId = creatorEmail.Id };

                            }
                            try
                            {
                                CSOMlistItem.SystemUpdate();
                                clientContext.ExecuteQuery();
                                _logger.LogInformation($"Createur et Modificateur OK");

                                //copiedDocumentSetListItem.SystemUpdate();
                                //await copiedDocumentSetItemToRename.SystemUpdateAsync();
                            }
                            catch (Exception e) { _logger.LogError(e.ToString()); }
                            
                            

                            

                            //managed metadata fields

                            var timeZone = TimeZoneInfo.FindSystemTimeZoneById("Romance Standard Time");
                            if (dateSouhaite != "")
                            {
                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
                                try
                                {

                                    DateTime ParsedDateSouhaite;
                                    try
                                    {
                                        // Try parsing with French culture
                                        ParsedDateSouhaite = DateTime.Parse(dateSouhaite, new CultureInfo("fr-FR"));
                                    }
                                    catch (FormatException)
                                    {
                                        // If the French parsing fails, try English culture
                                        ParsedDateSouhaite = DateTime.Parse(dateSouhaite, new CultureInfo("en-US"));
                                    }



                                    _logger.LogInformation(ParsedDateSouhaite.ToString(new CultureInfo("fr-FR")));
                                    var formValues3 = new List<ListItemFormUpdateValue>
                                                                                     {
                                                                            new ListItemFormUpdateValue() { FieldName = "Date_x0020_Souhait_x00e9_e" , FieldValue = ParsedDateSouhaite.ToString(new CultureInfo("fr-FR"))}

                                                                                     };
                                    CSOMlistItem.ValidateUpdateListItem(formValues3, false, "", false, false, "");
                                    clientContext.ExecuteQuery();
                                    _logger.LogInformation($"Date souhaite OK");
                                }
                                catch (Exception e) { _logger.LogError(e.ToString()); }

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
                                _logger.LogInformation($"Cree et modifie OK");

                            }
                            catch (Exception e) { _logger.LogError(e.ToString()); }


                            

                            if (!string.IsNullOrEmpty(specialiteMetier) && ListName == "Production de documents")
                            {
                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
                                var formValues1 = new List<ListItemFormUpdateValue>
                                                                                    {
                                                                            new ListItemFormUpdateValue() { FieldName = "Sp_x00e9_cialit_x00e9__x002f_m_x00e9_tier" , FieldValue = specialiteMetier}
                                                                                    };

                                try
                                {

                                    CSOMlistItem.ValidateUpdateListItem(formValues1, false, "", false, false, "");
                                    clientContext.ExecuteQuery();
                                    _logger.LogInformation($"Specialite metier OK pour prod de documents");
                                }
                                catch (Exception e) { _logger.LogError(e.ToString()); }

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
                                    _logger.LogInformation($"Specialite metier OK pour non prod de documents");
                                }
                                catch (Exception e) { _logger.LogError(e.ToString()); }

                            }

                            if (!string.IsNullOrEmpty(typeDeDocument))
                            {
                                CSOMlistItem.RefreshLoad();
                                clientContext.ExecuteQuery();
                                var formValues2 = new List<ListItemFormUpdateValue>
                                                                                     {
                                                                            new ListItemFormUpdateValue() { FieldName = "Typededocument" , FieldValue = typeDeDocument}
                                                                                     };
                                try
                                {
                                    CSOMlistItem.ValidateUpdateListItem(formValues2, false, "", false, false, "");
                                    clientContext.ExecuteQuery();
                                    _logger.LogInformation($"Type de document OK");

                                }
                                catch (Exception e) { _logger.LogError(e.ToString()); }



                            }
                            
                            //Only non system update update. To increment version number and update the modified timestamp
                            copiedDocumentSetListItem.Load();
                            await copiedDocumentSetListItem.UpdateAsync();
                           // copiedDocumentSetItemToRename.SystemUpdate();
                            _logger.LogInformation($"Document Set copied and renamed to: {title}. Metadata successfully updated");
                            //Prod statut webhook
                          // item["Statutwebhook"] = "Crée";
                            //PPR statut webhook
                            item["Statut_x0020_webhook"] = "Crée";
                            await item.UpdateAsync();
                            _logger.LogInformation($"Item status updated to created in og list");
                            if(isModified == true) { targetUrl2 = targetUrl2.Substring(0, targetUrl2.LastIndexOf('/') + 1) + title; }
                            return new  {targetURL=targetUrl2, titleDossier = title, modecreation = creationMode, OperationIdentifier = operationIdentifier, authorEmail = UserEmail, isModified = isModified };

                        }
                        else
                        {
                            _logger.LogError($"Failed to find the copied Document Set at {targetUrl2}.");
                            return null;
                        }
                    }
                    else
                    {
                        _logger.LogError($"Template NOT FOUND");
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
        public static async Task SaveLatestChangeTokenAsync(IChangeToken changeToken, string ressource, string blobconnection, string container, Microsoft.Extensions.Logging.ILogger _logger)
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
                _logger.LogError("Error saving change token");
                Console.WriteLine($"Error: {ex}");

            }

        }
        public static string FormatEmails(string emailList, string prefix, Microsoft.Extensions.Logging.ILogger _logger)
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
        public static string GetItemIdsBySiteTitles(ClientContext context, string listTitle, string values, string columnName, Microsoft.Extensions.Logging.ILogger _logger)
        {
            //This helper method gets the id of an item in a list given a column value (column to be specified)
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
                        
                        _logger.LogInformation($"List item with title '{value}' not found.");
                        concatenatedIds.Add("");
                    }
                }

                return string.Join(";", concatenatedIds);
            }
            catch { return string.Empty; }
        }

        public static string GetPaysIDFromProjet(ClientContext context, string listTitle, int? ProjetID, Microsoft.Extensions.Logging.ILogger _logger)
        {
            //This helper method gets the id of an item in a list given a column value (column to be specified)
            if (string.IsNullOrEmpty(ProjetID.ToString()))
                return string.Empty;
            try
            {

               

                var list = context.Web.Lists.GetByTitle(listTitle);
                var Projet = list.GetItemById(ProjetID.ToString());
                context.Load(Projet);
                context.ExecuteQuery();


                
                var pays = Projet.FieldValues["Pays"].ToString();

                _logger.LogInformation($"Pays from ProjetID");

                return string.IsNullOrEmpty(pays)? string.Empty : pays  ;
            }
            catch { return string.Empty; }
        }

        public static async Task<string> GetLatestChangeTokenAsync(string ressource, string blobconnection, string container, Microsoft.Extensions.Logging.ILogger _logger)
        {
            {
                BlobContainerClient containerClient = new BlobContainerClient(blobconnection, container);

                // Ensure the container exists
                await containerClient.CreateIfNotExistsAsync();

                // Verify if the container exists
                bool containerExists = await containerClient.ExistsAsync();
                if (!containerExists)
                {
                    _logger.LogError($"Blob container '{container}' does not exist and could not be created.");
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

        public static void Mailer(ClientContext context, string mailSubjectTemplate, string mailBodyTemplate, string authorEmail, string siteTitle, string siteURL, float storagePercentageAllowed, Microsoft.Extensions.Logging.ILogger _logger)
        {

            var MailSubject = mailSubjectTemplate.Replace("[SITE]", siteTitle);
            var MailBody = mailBodyTemplate.Replace("[PERCENTAGE]", storagePercentageAllowed.ToString() + "%").Replace("[SITE]", $"<a href='{siteURL}'>{siteTitle}</a>");
            var emailProperties = new EmailProperties
            {
                To = new[] { authorEmail } ,
                Subject = MailSubject,
                Body = MailBody,
            };

            try { Utility.SendEmail(context, emailProperties); }
            catch (Exception ex)
            {
                Console.WriteLine("Email send unsuccessful");
                Console.WriteLine(ex.Message);

            }

            Console.WriteLine("Email send successful");
            System.Threading.Thread.Sleep(2000);
        }
        public static  void Mailer2(ClientContext context,PnPContext pnpcontext, dynamic groupedItems, dynamic singleItems, Microsoft.Extensions.Logging.ILogger _logger)

        {
           
           
            // Process grouped items
            //var groupedItemLists = groupedItems.OfType<List<Item>>().ToList();
            foreach (dynamic itemList in groupedItems)
            {
                string authorEmailtoSend = null;
                // Build email body for the group
                var emailBodyBuilder = new System.Text.StringBuilder();
                emailBodyBuilder.AppendLine("<html><body>");
                emailBodyBuilder.AppendLine("Bonjour,");
                emailBodyBuilder.AppendLine("<br>");
                emailBodyBuilder.AppendLine("Les dossiers maitres suivants ont été créés : <br>");
                emailBodyBuilder.AppendLine( "<ul>");

                foreach ( var item in itemList)
                {
                
                    emailBodyBuilder.AppendLine($" <li><a href='{Uri.EscapeUriString("https://"+ pnpcontext.Web.Url.Host + item.targetURL)}'>{item.titleDossier}</a> </li>");
                   
                    authorEmailtoSend = item.authorEmail;
                }

                emailBodyBuilder.AppendLine("</ul>");
                emailBodyBuilder.AppendLine("Merci");
                emailBodyBuilder.AppendLine($"</body></html>");

                var MailBody = emailBodyBuilder.ToString();
                var emailProperties = new Microsoft.SharePoint.Client.Utilities.EmailProperties()
                {
                    To = new[] { authorEmailtoSend },
                    Subject = "Notification de création de dossier(s) maitre(s)",
                    Body = MailBody,
                    
                };
                // Send the email
                Utility.SendEmail(context,emailProperties);
                context.ExecuteQuery();
                _logger.LogInformation("Email Send Successful");
            }

            // Process single items
            foreach (dynamic singleItem in singleItems)
            {
                string authorEmailtoSend = null;
                // Build email body for the single item
                var emailBodyBuilder = new System.Text.StringBuilder();
                var modified = singleItem.isModified;
                emailBodyBuilder.AppendLine("<html><body>");
                emailBodyBuilder.AppendLine("Bonjour,");
                emailBodyBuilder.AppendLine("<br>");
                emailBodyBuilder.AppendLine($"Le dossier maitre suivant a été {(modified ? "modifié" : "créé")} : <br>");
                emailBodyBuilder.AppendLine("<ul>");
                emailBodyBuilder.AppendLine($"<li><a href='{Uri.EscapeUriString("https://" + pnpcontext.Web.Url.Host + singleItem.targetURL)}'>{singleItem.titleDossier}</a></li>");
                emailBodyBuilder.AppendLine("</ul>");
                emailBodyBuilder.AppendLine("Merci");
                emailBodyBuilder.AppendLine($"</body></html>");
                authorEmailtoSend = singleItem.authorEmail;
                var MailBody = emailBodyBuilder.ToString();

                var emailProperties = new EmailProperties
                {
                    To = new[] { authorEmailtoSend },
                    Subject = $"Notification de {(modified ? "modification" : "création")} de dossier(s) maitres(s)",
                    Body = MailBody,
                };
                // Send the email
                Utility.SendEmail(context, emailProperties);
                context.ExecuteQuery();
                _logger.LogInformation("Email Send Successful");
            }
        }

    }
}
