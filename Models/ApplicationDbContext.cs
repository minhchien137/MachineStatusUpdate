using Microsoft.EntityFrameworkCore;

namespace MachineStatusUpdate.Models
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
               : base(options)
        {
        }

        public DbSet<SVN_Equipment_Info_History> sVN_Equipment_Info_Histories { get; set; }
    }

}
