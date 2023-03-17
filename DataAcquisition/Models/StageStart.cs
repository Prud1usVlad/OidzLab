using System;
using System.Collections.Generic;

namespace DataAcquisition.Models;

public partial class StageStart
{
    public Guid Id { get; set; }

    public int? Stage { get; set; }

    public virtual Event IdNavigation { get; set; } = null!;
}
