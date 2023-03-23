namespace DataAcquisition.Models;

public partial class CurrencyPurchase
{
    public Guid Id { get; set; }

    public string? PackName { get; set; }

    public decimal? Price { get; set; }

    public int? Currency { get; set; }

    public virtual Event IdNavigation { get; set; } = null!;
}
