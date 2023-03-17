using System;
using System.Collections.Generic;

namespace DataAcquisition.Models;

public partial class StageEnd
{
    public Guid Id { get; set; }

    public int? Stage { get; set; }

    public bool? Win { get; set; }

    public int? Time { get; set; }

    public int? Currency { get; set; }

    public virtual Event IdNavigation { get; set; } = null!;
}
