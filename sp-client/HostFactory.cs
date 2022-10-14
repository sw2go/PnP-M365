using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace sp_client
{
    internal class HostFactory
    {
        /// <summary>
        /// read Config from appsettings.json
        /// </summary>
        /// <returns></returns>
        public static IHost CreateHostFromAppSettings()
        {
            var host = Host.CreateDefaultBuilder()
            // Configure logging
            .ConfigureServices((hostingContext, services) =>
            {
                // Add the PnP Core SDK library services
                services.AddPnPCore();
                // Add the PnP Core SDK library services configuration from the appsettings.json file
                services.Configure<PnPCoreOptions>(hostingContext.Configuration.GetSection("PnPCore"));
                // Add the PnP Core SDK Authentication Providers
                services.AddPnPCoreAuthentication();
                // Add the PnP Core SDK Authentication Providers configuration from the appsettings.json file
                services.Configure<PnPCoreAuthenticationOptions>(hostingContext.Configuration.GetSection("PnPCore"));
            })
            // Let the builder know we're running in a console
            .UseConsoleLifetime()
            // Add services to the container
            .Build();

            return host;
        }

        public static IHost CreateHost(string clientId, string tenantId)
        {
            var host = Host.CreateDefaultBuilder()
            // Configure logging
            .ConfigureServices((hostingContext, services) =>
            {
                // Add the PnP Core SDK library
                services.AddPnPCore(
                    options => {
                        options.PnPContext.GraphFirst = true;
                        options.HttpRequests.UserAgent = "ISV|Contoso|ProductX";

                        options.Sites.Add("SiteConfig1", new PnPCoreSiteOptions
                        {
                            SiteUrl = "https://contoso.sharepoint.com/sites/1008"
                        });
                    });

                services.AddPnPCoreAuthentication(
                    options => {

                        // Configure an Authentication Provider relying on the interactive authentication
                        options.Credentials.Configurations.Add("interactiveAuth",
                            new PnPCoreAuthenticationCredentialConfigurationOptions
                            {
                                ClientId = clientId,
                                TenantId = tenantId,
                                Interactive = new PnPCoreAuthenticationInteractiveOptions
                                {
                                    RedirectUri = new Uri("http://localhost")
                                }
                            });

                        // Configure an Authentication Provider relying on Windows Credential Manager
                        options.Credentials.Configurations.Add("certificateAuth",
                            new PnPCoreAuthenticationCredentialConfigurationOptions
                            {
                                ClientId = clientId,
                                TenantId = tenantId,
                                X509Certificate = new PnPCoreAuthenticationX509CertificateOptions
                                {
                                    StoreName = StoreName.My,
                                    StoreLocation = StoreLocation.CurrentUser,
                                    Thumbprint = "D7443BAB622CACF7653D7E97FB8D21639CE4A6EE"
                                }
                            });

                        // Configure an Authentication Provider relying on Windows Credential Manager
                        options.Credentials.Configurations.Add("credManagerAuth",
                            new PnPCoreAuthenticationCredentialConfigurationOptions
                            {
                                ClientId = clientId,
                                TenantId = tenantId,
                                CredentialManager = new PnPCoreAuthenticationCredentialManagerOptions
                                {   
                                    CredentialManagerName = "entryNotWorkingWithMFA"
                                }
                            });

                        // Configure the default authentication provider
                        options.Credentials.DefaultConfiguration = "certificateAuth";

                        // Map the site defined in AddPnPCore with the 
                        // Authentication Provider configured in this action
                        options.Sites.Add("SiteConfig1",
                            new PnPCoreAuthenticationSiteOptions
                            {
                                AuthenticationProviderName = "interactiveAuth"
                            });
                    });
            })
            // Let the builder know we're running in a console
            .UseConsoleLifetime()
            // Add services to the container
            .Build();

            return host;
        }

    }
}
