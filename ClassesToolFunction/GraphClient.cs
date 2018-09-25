using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClassesToolFunction;
using Microsoft.Azure;
using Microsoft.Extensions.Configuration.EnvironmentVariables;

namespace ClassesToolGraph
{
    public static class GraphClient
    {
        private static readonly string AuthEndpoint = "https://login.windows.net/";
        const string ResourceUrl = "https://graph.microsoft.com";
        static readonly string AppID = Enviroments.GetEnvironmentVariable("AppID");
        static readonly string Secret = Enviroments.GetEnvironmentVariable("Secret");
        static readonly string Tenant = Enviroments.GetEnvironmentVariable("Tenant");

        public static GraphServiceClient Create()
        {
            return new GraphServiceClient("https://graph.microsoft.com/v1.0",
                          new DelegateAuthenticationProvider(async (request) =>
                          {
                              string authority = AuthEndpoint + Tenant;
                              var authenticationContext = new AuthenticationContext(authority, false);
                              var clientCred = new ClientCredential(AppID, Secret);
                              AuthenticationResult authResult = await authenticationContext.AcquireTokenAsync(ResourceUrl, clientCred);
                              request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                          }));

        }
        public static async Task<Organization> GetOrganization(this GraphServiceClient graphClient)
        {
            var orgResult = await graphClient.Organization.Request().GetAsync();
            var res = orgResult.NextPageRequest;
            var orgList = orgResult.CurrentPage.ToList();

            Organization tenant = null;
            if (orgList.Count > 0)
            {
                tenant = orgList.First();
            }
            return tenant;
        }


        public static async Task<bool> SendMailAsync(this GraphServiceClient graphClient, Message msg, string user)
        {
            try
            {
                await graphClient.Users[user].SendMail(msg, true).Request().PostAsync();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        public static async Task<bool> RemoveMailAsync(this GraphServiceClient graphClient, Message msg, string user)
        {
            try
            {
                await graphClient.Users[user].Messages[msg.Id].Request().DeleteAsync();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        public static async Task<List<Message>> GetMailAsync(this GraphServiceClient graphClient, string user)
        {
            List<Message> result = new List<Message>();
            try
            {
                var list = await graphClient.Users[user].MailFolders.Inbox.Messages.Request().GetAsync();
                result.AddRange(list.CurrentPage);
                while (list.NextPageRequest != null)
                {
                    list = await list.NextPageRequest.GetAsync().ConfigureAwait(false);
                    result.AddRange(list.CurrentPage);
                }
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return result;
            }
        }

        public static async Task<bool> AddCalendarEntryAsync(this GraphServiceClient graphClient, Event ev, string user)
        {
            try
            {
                await graphClient.Users[user].Calendar.Events.Request().AddAsync(ev);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        public static async Task<bool> RemoveCalendarEntryAsync(this GraphServiceClient graphClient, Event ev, string user)
        {
            try
            {
                await graphClient.Users[user].Calendar.Events[ev.Id].Request().DeleteAsync();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return false;
            }
        }

        public static async Task<List<Event>> GetCalendarAsync(this GraphServiceClient graphClient, string user)
        {
            List<Event> result = new List<Event>();
            try
            {
                // https://developer.microsoft.com/en-us/graph/docs/concepts/query_parameters#filter-parameter
                //String filter = "start/dateTime gt '2018-09-11T01:00:00.000Z' and end/dateTime lt '2018-09-20T00:00:00.000Z'";

                String filter = String.Format("start/dateTime gt '{0}T00:00' and end/dateTime lt '{1}T00:00'",
                    DateTime.Now.AddDays(-5).ToString("yyyy-MM-dd"),
                    DateTime.Now.AddDays(6).ToString("yyyy-MM-dd"));

                var list = await graphClient.Users[user].Calendar.Events.Request().Filter(filter).GetAsync();
                result.AddRange(list.CurrentPage);
                while (list.NextPageRequest != null)
                {
                    list = await list.NextPageRequest.GetAsync().ConfigureAwait(false);
                    result.AddRange(list.CurrentPage);
                }
                return result;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return result;
            }
        }

    }
}
