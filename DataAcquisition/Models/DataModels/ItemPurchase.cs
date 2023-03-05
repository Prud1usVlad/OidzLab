using System;
using System.Collections.Generic;

namespace DataAcquisition.Models.DataModels;

public partial class ItemPurchase
{
    public Guid Id { get; set; }

    public string? ItemName { get; set; }

    public long? Price { get; set; }

    public virtual Event IdNavigation { get; set; } = null!;
}
