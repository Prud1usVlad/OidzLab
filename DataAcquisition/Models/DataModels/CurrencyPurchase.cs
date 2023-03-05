using System;
using System.Collections.Generic;

namespace DataAcquisition.Models.DataModels;

public partial class CurrencyPurchase
{
    public Guid Id { get; set; }

    public string? PackName { get; set; }

    public decimal? Price { get; set; }

    public long? Currency { get; set; }

    public virtual Event IdNavigation { get; set; } = null!;
}
