using System;
using System.Collections.Generic;

namespace DataAcquisition.Models.DataModels;

public partial class StageEnd
{
    public Guid Id { get; set; }

    public long? Stage { get; set; }

    public bool? Win { get; set; }

    public long? Time { get; set; }

    public long? Currency { get; set; }

    public virtual Event IdNavigation { get; set; } = null!;
}
