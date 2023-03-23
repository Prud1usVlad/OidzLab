namespace DataAcquisition.Models;

public partial class ItemPurchase
{
    public Guid Id { get; set; }

    public string? ItemName { get; set; }

    public int? Price { get; set; }

    public virtual Event IdNavigation { get; set; } = null!;
}
