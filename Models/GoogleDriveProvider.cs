using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using gAcss.Service;
using Google.Apis.Drive.v3;

namespace gAcss.Models
{
    public class GoogleDriveProvider : IDriveService
    {
        private readonly DriveService _service;

        public GoogleDriveProvider(DriveService service) => _service = service;

        public async Task<List<DriveFileEntry>> GetAllFilesAsync(CancellationToken ct = default)
        {
            var result = new List<DriveFileEntry>();
            string? pageToken = null;

            do
            {
                var request = _service.Files.List();
                request.Fields = "nextPageToken, files(id, name, size, modifiedTime, mimeType, owners, permissions(emailAddress, role, type), webViewLink)";
                request.PageSize = 1000;
                request.PageToken = pageToken;

                var response = await request.ExecuteAsync(ct);
                result.AddRange(MapToEntries(response.Files));
                pageToken = response.NextPageToken;
            } while (pageToken != null);

            return result;
        }

        public async Task<List<DriveFileEntry>> GetFolderFilesAsync(string rootFolderId, CancellationToken ct = default)
        {
            var result = new List<DriveFileEntry>();
            var foldersToProcess = new Stack<string>();

            foldersToProcess.Push(rootFolderId);

            while (foldersToProcess.Count > 0)
            {
                ct.ThrowIfCancellationRequested();

                string currentFolderId = foldersToProcess.Pop();
                string? pageToken = null;

                do
                {
                    var request = _service.Files.List();
                    request.Q = $"'{currentFolderId}' in parents and trashed = false";
                    request.Fields = "nextPageToken, files(id, name, size, modifiedTime, mimeType, owners, permissions(emailAddress, role, type), webViewLink)";
                    request.PageSize = 1000;
                    request.PageToken = pageToken;

                    var response = await request.ExecuteAsync(ct);
                    if (response.Files == null) break;

                    foreach (var file in response.Files)
                    {
                        if (file.MimeType == "application/vnd.google-apps.folder")
                        {
                            foldersToProcess.Push(file.Id);
                        }

                        result.AddRange(MapToEntries(new[] { file }));
                    }

                    pageToken = response.NextPageToken;
                } while (pageToken != null);
            }

            return result;
        }

        private IEnumerable<DriveFileEntry> MapToEntries(IEnumerable<Google.Apis.Drive.v3.Data.File> files)
        {
            foreach (var file in files)
            {
                var ownerEmail = file.Owners?.FirstOrDefault()?.EmailAddress;

                if (file.Permissions?.Any() == true)
                {
                    foreach (var p in file.Permissions)
                        yield return new DriveFileEntry(
                            file.Name,
                            file.WebViewLink,
                            p.EmailAddress,
                            p.Role,
                            p.Type,
                            file.Size,
                            file.ModifiedTimeDateTimeOffset?.DateTime,
                            file.MimeType,
                            ownerEmail);
                }
                else
                {
                    yield return new DriveFileEntry(
                        file.Name, file.WebViewLink, "N/A", "N/A", "N/A",
                        file.Size, file.ModifiedTimeDateTimeOffset?.DateTime, file.MimeType, ownerEmail);
                }
            }
        }
    }
}
