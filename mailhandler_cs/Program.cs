using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Kiota.Abstractions;
using Microsoft.Graph.Users.Item.MailFolders.Item.Messages;
using Microsoft.Graph.Users.Item.MailFolders;
using Azure.Identity;

namespace mailhandler_cs
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Konfiguration laden
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            var settings = config.GetSection("MailSettings");
            string? emailAddress = settings["EmailAddress"];
            string? clientId = settings["ClientId"];
            string? tenantId = settings["TenantId"];
            string? clientSecret = settings["ClientSecret"];
            string mailDir = settings["MailDir"] ?? "mails";
            bool includeFolders = bool.Parse(settings["IncludeFolders"] ?? "true");
            bool includeArchive = bool.Parse(settings["IncludeArchive"] ?? "true");
            int chunkSize = int.Parse(settings["ChunkSize"] ?? "50");
            bool loadAllEmails = bool.Parse(settings["LoadAllEmails"] ?? "true");
            int maxEmailsPerFolder = int.Parse(settings["MaxEmailsPerFolder"] ?? "0");
            int daysBack = int.Parse(settings["DaysBack"] ?? "30");
            int maxEmails = int.Parse(settings["MaxEmails"] ?? "100");

            if (string.IsNullOrWhiteSpace(emailAddress)) throw new Exception("EmailAddress fehlt in der Konfiguration!");
            if (string.IsNullOrWhiteSpace(clientId)) throw new Exception("ClientId fehlt in der Konfiguration!");
            if (string.IsNullOrWhiteSpace(tenantId)) throw new Exception("TenantId fehlt in der Konfiguration!");
            if (string.IsNullOrWhiteSpace(clientSecret)) throw new Exception("ClientSecret fehlt in der Konfiguration!");

            Directory.CreateDirectory(mailDir);
            Directory.CreateDirectory(Path.Combine(mailDir, "pdf"));

            // Authentifizierung: Credential erstellen
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            var graphClient = new GraphServiceClient(credential);

            Console.WriteLine($"Starte E-Mail-Download für {emailAddress}...");

            var sinceDate = DateTime.UtcNow.AddDays(-daysBack);
            int totalDownloaded = 0;

            // 1. Posteingang
            totalDownloaded += await DownloadFromFolder(graphClient, emailAddress, "Inbox", mailDir, sinceDate, chunkSize, loadAllEmails, maxEmails, maxEmailsPerFolder);

            // 2. Unterordner
            if (includeFolders)
            {
                var folders = await GetAllFolders(graphClient, emailAddress);
                foreach (var folder in folders)
                {
                    if (folder.DisplayName != null && folder.DisplayName.ToLower() != "inbox" && folder.DisplayName.ToLower() != "archive")
                    {
                        totalDownloaded += await DownloadFromFolder(graphClient, emailAddress, folder.Id, mailDir, sinceDate, chunkSize, loadAllEmails, maxEmails, maxEmailsPerFolder, folder.DisplayName);
                    }
                }
            }

            // 3. Archiv
            if (includeArchive)
            {
                totalDownloaded += await DownloadFromFolder(graphClient, emailAddress, "Archive", mailDir, sinceDate, chunkSize, loadAllEmails, maxEmails, maxEmailsPerFolder);
            }

            Console.WriteLine($"Fertig. Insgesamt {totalDownloaded} E-Mails heruntergeladen.");
        }

        static async Task<int> DownloadFromFolder(GraphServiceClient graphClient, string emailAddress, string folderIdOrName, string mailDir, DateTime sinceDate, int chunkSize, bool loadAllEmails, int maxEmails, int maxEmailsPerFolder, string displayName = "")
        {
            int downloaded = 0;
            string folderName = !string.IsNullOrEmpty(displayName) ? displayName : folderIdOrName;
            var filter = $"receivedDateTime ge {sinceDate:yyyy-MM-ddTHH:mm:ssZ}";

            var messagesPage = await graphClient.Users[emailAddress].MailFolders[folderIdOrName].Messages.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Filter = filter;
                requestConfiguration.QueryParameters.Orderby = new[] { "receivedDateTime asc" };
                requestConfiguration.QueryParameters.Top = chunkSize;
                requestConfiguration.QueryParameters.Expand = new[] { "attachments" };
                requestConfiguration.QueryParameters.Select = new[] { "id", "subject", "from", "toRecipients", "receivedDateTime", "body", "bodyPreview", "attachments" };
            });

            while (messagesPage != null && messagesPage.Value != null && messagesPage.Value.Count > 0)
            {
                foreach (var mail in messagesPage.Value)
                {
                    string subject = SanitizeFilename(mail.Subject ?? "Kein_Betreff");
                    if (subject.Length > 50) subject = subject.Substring(0, 50);
                    string dateStr = mail.ReceivedDateTime?.ToString("yyyy-MM-dd-HH-mm-ss") ?? DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm-ss");
                    string filename = $"{dateStr}--[{folderName}]--{subject}.txt";
                    string filepath = Path.Combine(mailDir, filename);

                    string from = mail.From?.EmailAddress?.Address ?? "Unbekannt";
                    string to = mail.ToRecipients?.FirstOrDefault()?.EmailAddress?.Address ?? "Unbekannt";
                    string body = mail.Body?.Content ?? mail.BodyPreview ?? "Kein Inhalt verfügbar";

                    string emailText = $"Von: {from}\nAn: {to}\nDatum: {mail.ReceivedDateTime}\nBetreff: {mail.Subject}\nOrdner: {folderName}\n\n{body}";
                    await File.WriteAllTextAsync(filepath, emailText);

                    // PDF-Attachments speichern
                    if (mail.Attachments != null && mail.Attachments.Count > 0)
                    {
                        foreach (var att in mail.Attachments)
                        {
                            if (att is FileAttachment fileAtt && fileAtt.ContentType != null && fileAtt.ContentType.ToLower() == "application/pdf" && fileAtt.ContentBytes != null)
                            {
                                string pdfName = SanitizeFilename(fileAtt.Name ?? "Anhang.pdf");
                                if (!pdfName.ToLower().EndsWith(".pdf")) pdfName += ".pdf";
                                string pdfFilename = $"{dateStr}--[{folderName}]--{pdfName}";
                                string pdfPath = Path.Combine(mailDir, "pdf", pdfFilename);
                                await File.WriteAllBytesAsync(pdfPath, fileAtt.ContentBytes);
                            }
                        }
                    }
                    downloaded++;
                    if (!loadAllEmails && downloaded >= maxEmails) return downloaded;
                    if (maxEmailsPerFolder > 0 && downloaded >= maxEmailsPerFolder) return downloaded;
                }
                if (messagesPage.OdataNextLink != null)
                {
                    var requestInfo = new RequestInformation
                    {
                        HttpMethod = Method.GET,
                        UrlTemplate = messagesPage.OdataNextLink
                    };
                    messagesPage = await graphClient.RequestAdapter.SendAsync<Microsoft.Graph.Models.MessageCollectionResponse>(requestInfo, Microsoft.Graph.Models.MessageCollectionResponse.CreateFromDiscriminatorValue);
                }
                else
                {
                    messagesPage = null;
                }
            }
            return downloaded;
        }

        static async Task<List<MailFolder>> GetAllFolders(GraphServiceClient graphClient, string emailAddress)
        {
            var folders = new List<MailFolder>();
            var page = await graphClient.Users[emailAddress].MailFolders.GetAsync();
            while (page != null && page.Value != null && page.Value.Count > 0)
            {
                folders.AddRange(page.Value);
                if (page.OdataNextLink != null)
                {
                    var requestInfo = new RequestInformation
                    {
                        HttpMethod = Method.GET,
                        UrlTemplate = page.OdataNextLink
                    };
                    page = await graphClient.RequestAdapter.SendAsync<Microsoft.Graph.Models.MailFolderCollectionResponse>(requestInfo, Microsoft.Graph.Models.MailFolderCollectionResponse.CreateFromDiscriminatorValue);
                }
                else
                {
                    page = null;
                }
            }
            return folders;
        }

        static string SanitizeFilename(string filename)
        {
            if (filename == null) return "Unbenannt";
            foreach (var c in Path.GetInvalidFileNameChars())
                filename = filename.Replace(c, '_');
            while (filename.Contains("__")) filename = filename.Replace("__", "_");
            return filename.Trim();
        }
    }
}
