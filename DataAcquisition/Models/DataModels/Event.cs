using System;
using System.Collections.Generic;

namespace DataAcquisition.Models.DataModels;

public partial class Event
{
    public Guid Id { get; set; }

    public DateOnly? Date { get; set; }

    public Guid? UserId { get; set; }

    public long? Type { get; set; }

    public virtual CurrencyPurchase? CurrencyPurchase { get; set; }

    public virtual ItemPurchase? ItemPurchase { get; set; }

    public virtual StageEnd? StageEnd { get; set; }

    public virtual StageStart? StageStart { get; set; }

    public virtual User? User { get; set; }
}
