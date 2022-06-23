using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Configuration;
using Azure.Identity;
using Microsoft.Graph;
//using System.Diagnostics;

namespace Managers
{
    public class OnedriveManager
    {
        private static readonly string[] _scopes;
        private static readonly string _tenantId;
        private static readonly string _clientId;
        private static readonly string _clientSecret;
        private static readonly string _driveId;
        private static readonly string _siteId;
        private static readonly TokenCredentialOptions _options;
        private static readonly ClientSecretCredential _clientSecretCredential;
        private static readonly GraphServiceClient _graphClient;

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static OnedriveManager()
        {
            try
            {
                //Debug.WriteLine("-----------------------------TEST-------------------------");
                _scopes = new[] { WebConfigurationManager.AppSettings["AZURE_PM_APP_SCOPES"] };

                // Multi-tenant apps can use "common",
                // single-tenant apps must use the tenant ID from the Azure portal
                _tenantId = WebConfigurationManager.AppSettings["AZURE_PM_TENANT_ID"];

                // Values from app registration
                _clientId = WebConfigurationManager.AppSettings["AZURE_PM_APP_ID"];
                _clientSecret = WebConfigurationManager.AppSettings["AZURE_PM_APP_SECRET"];

                // Sharepoint
                _siteId = WebConfigurationManager.AppSettings["AZURE_SHRP_SITE"];
                _driveId = WebConfigurationManager.AppSettings["AZURE_SHRP_DRIVE_ID"];

                _options = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };

                // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
                _clientSecretCredential = new ClientSecretCredential(
                _tenantId, _clientId, _clientSecret, _options);

                _graphClient = new GraphServiceClient(_clientSecretCredential, _scopes);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
            }
        }

        public static async Task<string> CreateFolder(string path)
        {

            if (String.IsNullOrEmpty(path))
            {
                return "";
            }
         
            var folder = new DriveItem
            {
                Name = path,
                Folder = new Folder()
            };

            try
            {
                var result = await _graphClient
                    //.Me
                    //.Drive
                    //.Root
                    .Sites[_siteId]
                    .Drives[_driveId]
                    .Root
                    .ItemWithPath("test1/test2")
                    .Children
                    .Request()
                    .AddAsync(folder);
                //.GetAsync();
                return result.ToString();
            }
            catch (ServiceException ex)
            {
                if (ex.IsMatch(GraphErrorCode.NameAlreadyExists.ToString()))
                {
                    System.Diagnostics.Debug.WriteLine(ex.Error.Code);
                }
                System.Diagnostics.Debug.WriteLine(ex.Message);
            };

            return "";
        }
    }
}
