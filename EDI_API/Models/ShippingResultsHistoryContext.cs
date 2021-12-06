using Microsoft.EntityFrameworkCore;

namespace EDI_API.Models
{
    public class ShippingResultsHistoryContext:DbContext
    {
        public ShippingResultsHistoryContext(DbContextOptions<ShippingResultsHistoryContext> options) : base(options)
        {

        }

        public DbSet<ShippingResultsHistory> Shipping_Results_History { get; set; }
    }
}
