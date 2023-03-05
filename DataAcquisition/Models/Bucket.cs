using System;
using System.Collections.Generic;

namespace DataAcquisition.Models;

public partial class Bucket
{
    public string Id { get; set; } = null!;

    public string Name { get; set; } = null!;

    public Guid? Owner { get; set; }

    public DateTime? CreatedAt { get; set; }

    public DateTime? UpdatedAt { get; set; }

    public bool? Public { get; set; }

    public bool? AvifAutodetection { get; set; }

    public long? FileSizeLimit { get; set; }

    public string[]? AllowedMimeTypes { get; set; }

    public virtual ICollection<Object> Objects { get; } = new List<Object>();

    public virtual User1? OwnerNavigation { get; set; }
}
