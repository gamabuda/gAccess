using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace gAcss.Models
{
    public record DriveFileEntry(
        string Name,
        string Link,
        string? Email,
        string? Role,
        string? Type,
        long? Size,           
        DateTime? Modified,  
        string? MimeType,     
        string? Owner         
    );
}
