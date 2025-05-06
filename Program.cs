using Microsoft.Azure.Functions.Worker; // Needed for ConfigureFunctionsWorkerDefaults()
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using FunctionApp3;

var host = new HostBuilder()
    .ConfigureAppConfiguration((context, config) =>
    {
        // Set base path and add configuration sources.
        config.SetBasePath(Environment.CurrentDirectory)
              .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
              .AddEnvironmentVariables();
    })
    // This extension method registers the Functions worker defaults for the isolated model.
    .ConfigureFunctionsWebApplication()
    .ConfigureServices((context, services) =>
    {
        // Register our custom settings options using our custom setup class.
        services.ConfigureOptions<FunctionAppSettingsSetup>();
        var anis = context.HostingEnvironment;
        // Register HttpClientFactory.
        services.AddHttpClient();

        // Register the SharePoint context factory.
        services.AddSingleton<ISharePointContextFactory, SharePointContextFactory>();
        services.AddSingleton<IServices, Services>();

        // Register Application Insights telemetry.
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        // Adjust logging filter options so that lower-level logs are captured.
        services.Configure<LoggerFilterOptions>(options =>
        {
            var rule = options.Rules.FirstOrDefault(r =>
                r.ProviderName == "Microsoft.Extensions.Logging.ApplicationInsights.ApplicationInsightsLoggerProvider");
            if (rule is not null)
            {
                options.Rules.Remove(rule);
            }
        });
    })
    .Build();

host.Run();


// ----------------------------
// Settings Classes & Setup
// ----------------------------

public class FunctionAppSettings
{
    // Environment-specific settings.
    public required string siteURL { get; set; }
    public required string AzureBlobStorageConnectionString { get; set; }
    public required string CSVDocumentLibraryTitle { get; set; }

    // Indiscriminate settings (same for both environments).
    public required string clientID { get; set; }
    public required string clientSecretID { get; set; }
    public required string BlobContainerCSV { get; set; }
    public required string BlobContainerList { get; set; }
    public required string CreationListName { get; set; }
    public required string AzureFunctionBaseURL { get; set; }

    public required string deploymentEnv { get; set; }
}

public class FunctionAppSettingsSetup : IConfigureOptions<FunctionAppSettings>
{
    private readonly IConfiguration _configuration;
    public FunctionAppSettingsSetup(IConfiguration configuration) {
        _configuration = configuration;
    }
    public void Configure(FunctionAppSettings options) {
        // Read DEPLOYMENT_ENV to choose between PPR and PROD (defaults to "PPR").
        string deploymentEnv = Environment.GetEnvironmentVariable("DEPLOYMENT_ENV") ?? "PPR";

        if (deploymentEnv.Equals("PROD", StringComparison.OrdinalIgnoreCase))
        {
            options.siteURL = _configuration["SharePoint_SiteUrl_GeostockPROD"]?.Trim().Replace("\"", "");
            options.AzureBlobStorageConnectionString = _configuration["AzureBlobStorageConnectionString"]?.Trim().Replace("\"", "");
            options.CSVDocumentLibraryTitle = _configuration["CSVDocumentLibraryTitle"]?.Trim().Replace("\"", "");
        }
        else
        {
            options.siteURL = _configuration["SharePoint_SiteUrl_GeostockPPR"]?.Trim().Replace("\"", "");
            options.AzureBlobStorageConnectionString = _configuration["AzureBlobStorageConnectionStringPPR"]?.Trim().Replace("\"", "");
            options.CSVDocumentLibraryTitle = _configuration["CSVDocumentLibraryTitlePPR"]?.Trim().Replace("\"", "");
        }

        // Indiscriminate settings.
        options.clientID = _configuration["SharePoint_ClientID"]?.Trim().Replace("\"", "");
        options.clientSecretID = _configuration["SharePoint_ClientSecretID"]?.Trim().Replace("\"", "");
        options.BlobContainerCSV = _configuration["BlobContainerCSV"]?.Trim().Replace("\"", "");
        options.BlobContainerList = _configuration["BlobContainerList"]?.Trim().Replace("\"", "");
        options.CreationListName = _configuration["CreationListName"]?.Trim().Replace("\"", "");
        options.AzureFunctionBaseURL = _configuration["AzureFunctionBaseURL"]?.Trim().Replace("\"", "");
        options.deploymentEnv = deploymentEnv;


    }
}

// -----------------------------
// SharePoint Context Factory
// -----------------------------
public interface ISharePointContextFactory
{
    /// <summary>
    /// Creates a SharePoint ClientContext using app-only authentication.
    /// </summary>
    Microsoft.SharePoint.Client.ClientContext CreateClientContext(string siteUrl, string clientId, string clientSecret);
}

public class SharePointContextFactory : ISharePointContextFactory
{
    public Microsoft.SharePoint.Client.ClientContext CreateClientContext(string siteUrl, string clientId, string clientSecret) {
        // Use PnP.Framework's AuthenticationManager to obtain a ClientContext.
        return new PnP.Framework.AuthenticationManager().GetACSAppOnlyContext(siteUrl, clientId, clientSecret);
    }
}