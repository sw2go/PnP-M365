using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;

namespace sp_client
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var host = HostFactory.CreateHostFromAppSettings();

            //var host = HostFactory.CreateHost("c74c39cc-176d-4fc7-8f61-a37670edf163", "17e07a74-67fd-4b09-bc5e-633cb11302ce");

            try
            {
                await host.StartAsync();

                using(var ctx = await Context.CreateContextAsync(host, "SiteConfig1"))
                {
                    var spFolder = await ctx.GetFolderAsync("/Shared Documents/General/Virtual-Machines");

                    foreach (var f in await spFolder.FindFilesAsync("ShortCuts.txt"))
                    {
                        await ctx.DownloadFileAsync(f, @"C:\@Tmp\@hampe");
                    }

                    await ctx.UploadFileAsync(@"C:\@Tmp\@hampe\sami.txt", spFolder);

                }

                              


                

            }
            catch (Exception ex)
            {

            }
        }


        public static async Task GetMessage(IHost host)
        {
            using (var scope = host.Services.CreateScope())
            {
                // Obtain a PnP Context factory
                var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();
                // Use the PnP Context factory to get a PnPContext for the given configuration
                using (var context = await pnpContextFactory.CreateAsync("SiteToWorkWith"))
                {
                        


                    var channel = await context.Team.Channels
                        .Where(i => i.DisplayName == "General") // "General" Channel only
                        .QueryProperties(p => p.Messages)       // Messages-Property dazuladen
                        .FirstOrDefaultAsync();

                    // ab C# 8.0 geht async foreach mit where-Filter und async AsAsyncEnumerable
                    // await foreach(var mes in channel.Messages.Where(m => m.Body.ContentType == PnP.Core.Model.Teams.ChatMessageContentType.Text).AsAsyncEnumerable())
                    // paged loading
                    foreach (var m in channel.Messages)
                    {
                        if (m.Body.ContentType == PnP.Core.Model.Teams.ChatMessageContentType.Text)
                        {
                            var text = m.Body.Content;

                            var diff = DateTimeOffset.UtcNow - m.LastModifiedDateTime.ToUniversalTime();
                            if (diff.TotalDays > 1) // bei älteren Meldungen -> Abbruch
                            {
                                break;
                            }
                        }
                    }
                }
            }

        }

        public static async Task AddMessage2(IHost host)
        {

            using (var scope = host.Services.CreateScope())
            {
                // Obtain a PnP Context factory
                var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();
                // Use the PnP Context factory to get a PnPContext for the given configuration
                using (var context = await pnpContextFactory.CreateAsync("SiteToWorkWith"))
                {
                    var channel = await context.Team.Channels
                        .Where(i => i.DisplayName == "General") // "General" Channel only
                        .QueryProperties(p => p.Messages)       // Messages-Property dazuladen
                        .FirstOrDefaultAsync();

                    var body = "PNP";

                    // Perform the add operation
                    await channel.Messages.AddAsync(body);

                }
            }

        }



        public static async Task AddMessage(IHost host)
        {

            using (var scope = host.Services.CreateScope())
            {
                // Obtain a PnP Context factory
                var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();
                // Use the PnP Context factory to get a PnPContext for the given configuration
                using (var context = await pnpContextFactory.CreateAsync("SiteToWorkWith"))
                {
                    var team = await context.Team.GetAsync(o => o.Channels);
                    var channel = team.Channels.AsRequested().FirstOrDefault(i => i.DisplayName == "General");

                    channel = await channel.GetAsync(o => o.Messages);
                    var chatMessages = channel.Messages;

                    var body = "PNP";

                    // Perform the add operation
                    await chatMessages.AddAsync(body);

                }
            }

        }




    }
}