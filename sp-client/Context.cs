using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace sp_client
{
    public class Context : IDisposable
    {
        private bool disposedValue;
        private PnPContext context;

        private Context(PnPContext context)
        {
            this.context = context;
        }

        public async static Task<Context> CreateContextAsync(IHost host, string siteConfigName)
        {
            using (var scope = host.Services.CreateScope())
            {
                // Obtain a PnP Context factory
                var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();

                // Use the PnP Context factory to get a PnPContext for the given configuration
                var pnpContext = await pnpContextFactory.CreateAsync(siteConfigName);

                return new Context(pnpContext);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="siteSubFolder"> default = Teams General Channel Folder</param>
        /// <returns></returns>
        public async Task<IFolder> GetFolderAsync(string siteSubFolder = "/Shared Documents/General")
        {
            string folderUrl = $"{context.Uri.PathAndQuery}{siteSubFolder}";
            return await context.Web.GetFolderByServerRelativeUrlAsync(folderUrl);
        }


        public async Task<IFile> UploadFileAsync(string localFilePath, IFolder remoteFolder)
        {
            var remoteFileName = Path.GetFileName(localFilePath);
            Stream uploadContentStream = System.IO.File.OpenRead(localFilePath);
            return await remoteFolder.Files.AddAsync(remoteFileName, uploadContentStream, true);
        }

        public async Task DownloadFileAsync(IFile remoteFile, string localFolder)
        {
            var localFileName = remoteFile.Name;
            Stream downloadedContentStream = await remoteFile.GetContentAsync(true);

            // Download the file bytes in 2MB chunks and immediately write them to a file on disk 
            // This approach avoids the file being fully loaded in the process memory
            var bufferSize = 2 * 1024 * 1024;  // 2 MB buffer
            using (var content = System.IO.File.Create(Path.Combine(localFolder, localFileName)))
            {
                var buffer = new byte[bufferSize];
                int read;
                while ((read = await downloadedContentStream.ReadAsync(buffer, 0, buffer.Length)) != 0)
                {
                    content.Write(buffer, 0, read);
                }
            }
        }

        public async Task AddMessage(string channelName, string text)
        {
            var channel = await context.Team.Channels
                .Where(i => i.DisplayName == channelName) // "General" Channel only
                .QueryProperties(p => p.Messages)       // Messages-Property dazuladen
                .FirstOrDefaultAsync();

            await channel.Messages.AddAsync(text);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {                    
                    this.context.Dispose();
                    // TODO: dispose other managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
