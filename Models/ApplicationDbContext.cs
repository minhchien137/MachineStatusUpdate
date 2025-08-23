using Microsoft.EntityFrameworkCore;

namespace MachineStatusUpdate.Models
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
               : base(options)
        {
        }

        public DbSet<SVN_Equipment_Info_History> SVN_Equipment_Info_History { get; set; }

        public DbSet<SVN_Equipment_Machine_Info> sVN_Equipment_Machine_Info { get; set; }
    }

}
