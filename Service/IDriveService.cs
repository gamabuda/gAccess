using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using gAcss.Models;

namespace gAcss.Service
{
    public interface IDriveService
    {
        Task<List<DriveFileEntry>> GetAllFilesAsync(CancellationToken ct = default);
        Task<List<DriveFileEntry>> GetFolderFilesAsync(string folderId, CancellationToken ct = default);
    }
}
