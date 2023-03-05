using System;
using System.Collections.Generic;

namespace DataAcquisition.Models.DataModels;

public partial class StageStart
{
    public Guid Id { get; set; }

    public long? Stage { get; set; }

    public virtual Event IdNavigation { get; set; } = null!;
}
