using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunctionApp3
{
    internal class Models
    {
        public class ResponseModel<T>
        {
            [JsonProperty(PropertyName = "value")]
            public List<T> Value { get; set; }
        }
        public class NotificationModel
        {
            [JsonProperty(PropertyName = "subscriptionId")]
            public string SubscriptionId { get; set; }

            [JsonProperty(PropertyName = "clientState")]
            public string ClientState { get; set; }

            [JsonProperty(PropertyName = "expirationDateTime")]
            public DateTime ExpirationDateTime { get; set; }

            [JsonProperty(PropertyName = "resource")]
            public string Resource { get; set; }

            [JsonProperty(PropertyName = "tenantId")]
            public string TenantId { get; set; }

            [JsonProperty(PropertyName = "siteUrl")]
            public string SiteUrl { get; set; }

            [JsonProperty(PropertyName = "webId")]
            public string WebId { get; set; }
        }
        public class JSONMetadataFieldInList
        {
            public string? ID { get; set; }
            public string? Title { get; set; }

          
            public string? CodeProjet { get; set; }
            public string? Source { get; set; }
            public string? NumeroDossier { get; set; }
            public string? Activite { get; set; }
            public string? Approbateurs { get; set; }
            public string? Emetteurs { get; set; }
            public string? Verificateurs { get; set; }
            public string? Secretaire { get; set; }
            public string? Chef { get; set; }
            public string? Bibliotheque { get; set; }
            public string? Langue { get; set; }
            public string? DateSouhaite { get; set; }
            public string? NombrePage { get; set; }
            public string? NombreAnnexes { get; set; }
            public string? Departement { get; set; }
            public string? DossierProjet { get; set; }
            public string? Client { get; set; }
            public string? Etat { get; set; }
            public string? IntervenantPAO { get; set; }
            public string? IntervenantTTX { get; set; }
            public string? PhaseEtude { get; set; }
            public string? SpecialiteMetier { get; set; }
            public string? TypeDeDocument { get; set; }
            public string? Societe { get; set; }
            public string? Revision { get; set; }
            public string? ObjetRevision { get; set; }
            public string? CodePays { get; set; }
            public string? IntituleProjet { get; set; }
            public string? MotsCles { get; set; }
            public string? Description { get; set; }
            public string? Site { get; set; }
        }
    }
}
