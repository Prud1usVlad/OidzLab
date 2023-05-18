using Microsoft.EntityFrameworkCore;

namespace DataAcquisition.Models;

public partial class OidzDbContext : DbContext
{
    public OidzDbContext()
    {
    }

    public OidzDbContext(DbContextOptions<OidzDbContext> options)
        : base(options)
    {
    }

    public virtual DbSet<CurrencyPurchase> CurrencyPurchases { get; set; }

    public virtual DbSet<Event> Events { get; set; }

    public virtual DbSet<ItemPurchase> ItemPurchases { get; set; }

    public virtual DbSet<StageEnd> StageEnds { get; set; }

    public virtual DbSet<StageStart> StageStarts { get; set; }

    public virtual DbSet<User> Users { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Data Source=178.165.88.5\\DEV;Initial Catalog=OidzDb;User ID=user1;Password=1111;TrustServerCertificate=True");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<CurrencyPurchase>(entity =>
        {
            entity.ToTable("CurrencyPurchase");

            entity.Property(e => e.Id).ValueGeneratedOnAdd();
            entity.Property(e => e.PackName).HasMaxLength(200);
            entity.Property(e => e.Price).HasColumnType("decimal(18, 0)");

            entity.HasOne(d => d.IdNavigation).WithOne(p => p.CurrencyPurchase)
                .HasForeignKey<CurrencyPurchase>(d => d.Id)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_CurrencyPurchase_CurrencyPurchase");
        });

        modelBuilder.Entity<Event>(entity =>
        {
            entity.ToTable("Event");

            entity.Property(e => e.Date).HasColumnType("date");

            entity.HasOne(d => d.User).WithMany(p => p.Events)
                .HasForeignKey(d => d.UserId)
                .HasConstraintName("FK_Event_User");
        });

        modelBuilder.Entity<ItemPurchase>(entity =>
        {
            entity.ToTable("ItemPurchase");

            entity.Property(e => e.Id).ValueGeneratedNever();
            entity.Property(e => e.ItemName).HasMaxLength(200);

            entity.HasOne(d => d.IdNavigation).WithOne(p => p.ItemPurchase)
                .HasForeignKey<ItemPurchase>(d => d.Id)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_ItemPurchase_Event");
        });

        modelBuilder.Entity<StageEnd>(entity =>
        {
            entity.ToTable("StageEnd");

            entity.Property(e => e.Id).ValueGeneratedNever();

            entity.HasOne(d => d.IdNavigation).WithOne(p => p.StageEnd)
                .HasForeignKey<StageEnd>(d => d.Id)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_StageEnd_Event");
        });

        modelBuilder.Entity<StageStart>(entity =>
        {
            entity.ToTable("StageStart");

            entity.Property(e => e.Id).ValueGeneratedNever();

            entity.HasOne(d => d.IdNavigation).WithOne(p => p.StageStart)
                .HasForeignKey<StageStart>(d => d.Id)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_StageStart_Event");
        });

        modelBuilder.Entity<User>(entity =>
        {
            entity.ToTable("User");

            entity.Property(e => e.Id).HasDefaultValueSql("(newid())");
            entity.Property(e => e.Country).HasMaxLength(200);
            entity.Property(e => e.Gender).HasMaxLength(50);
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
