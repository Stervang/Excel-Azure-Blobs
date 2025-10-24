using Azure;
using Azure.Core;
using Azure.Core.Pipeline;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using ExcelDna.Integration;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RibbonStartCopy
{
    internal static class AzureControl
    {
        private static readonly HttpClient handler = new HttpClient();
        private static readonly BlobClientOptions options = new BlobClientOptions
        {
            Transport = new Azure.Core.Pipeline.HttpClientTransport(
        new HttpClient {Timeout = TimeSpan.FromMinutes(5) }),
            Retry =
                        {
                            MaxRetries = 3,
                            Mode = RetryMode.Exponential,
                            Delay = TimeSpan.FromSeconds(5),
                            MaxDelay = TimeSpan.FromSeconds(20),
                            NetworkTimeout = TimeSpan.FromSeconds(120) // adjust as needed

                        }
        };

        // Static shared HttpClient to avoid socket teardown issues
        private static readonly HttpClient SharedHttpClient = new HttpClient();

        // Shared BlobClientOptions using the shared HttpClient
        private static readonly BlobClientOptions SharedOptions = new BlobClientOptions
        {
            Transport = new HttpClientTransport(SharedHttpClient)
        };
        public static bool GetAzureData(string FileID, string location)
        {
            try
            {
                return true;//Task.Run(() => GetAzureDataAsync(FileID, location)).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Download failed: " + ex.Message);
                return false;
            }
        }

        [ExcelCommand]
        public static void GetAzureDataAsync(string FileID, string location)
        {
            int attempts = 0;
            const int maxAttempts = 3;

            System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

            Thread t = new Thread(() =>
            {
                string connectionString = "UseDevelopmentStorage=true";

                BlobDownloadStreamingResult downloadResult = null;

                //var blobServiceClient = new BlobServiceClient(connectionString, SharedOptions);
                string containerName = "user-123";
                //var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                //var blobClient = containerClient.GetBlobClient(FileID);

                var client = new BlobServiceClient(connectionString);
                var containers = client.GetBlobContainerClient(containerName);

                var blobClient = containers.GetBlobClient(FileID);

                //if (!await blobClient.ExistsAsync())
                //{
                //    //ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show("Blob does not exist."));
                //    return string.Empty;
                //}

                downloadResult = blobClient.DownloadStreaming();
                using var contentStream = downloadResult.Content;
                using var fs = new FileStream(location, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
                contentStream.CopyTo(fs);
                //fs.FlushAsync();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //Task.Delay(300).ConfigureAwait(false);

                //attempts++;
                //await Task.Delay(1000); // Optional back-off
                //return string.Empty;
            });

            t.Start();

        }

        [ExcelCommand]
        public static void GetAzureDataAsyncOmni(string FileID, string location)
        {
            int attempts = 0;
            const int maxAttempts = 3;

            //System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            System.Net.ServicePointManager.ServerCertificateValidationCallback = (sender, cert, chain, errors) => true;
            Thread t = new Thread(() =>
            {
                string connectionString = "DefaultEndpointsProtocol=https;AccountName=excelblobdemo;AccountKey=b4BK8VWQ+L3CYalKKaov8x0Oh7UVm22IZOq2mKTpeBxhDzg5sTZd+U4YyNWGKKzaCzLm73CXMC8++AStLaxQzQ==;EndpointSuffix=core.windows.net";

                BlobDownloadStreamingResult downloadResult = null;

                //var blobServiceClient = new BlobServiceClient(connectionString, SharedOptions);
                string containerName = "models";
                //var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                //var blobClient = containerClient.GetBlobClient(FileID);

                var client = new BlobServiceClient(connectionString, options);
                var containers = client.GetBlobContainerClient(containerName);

                var blobClient = containers.GetBlobClient(FileID);

                //if (!await blobClient.ExistsAsync())
                //{
                //    //ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show("Blob does not exist."));
                //    return string.Empty;
                //}

                downloadResult = blobClient.DownloadStreaming();
                using var contentStream = downloadResult.Content;
                using var fs = new FileStream(location, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
                contentStream.CopyTo(fs);
                //fs.FlushAsync();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //Task.Delay(300).ConfigureAwait(false);

                //attempts++;
                //await Task.Delay(1000); // Optional back-off
                //return string.Empty;
            });

            t.Start();

        }
    }
}
