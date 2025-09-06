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

        public DbSet<SVN_Equipment_Info_History_Test> SVN_Equipment_Info_History_Test { get; set; }

        public DbSet<SVN_Equipment_Machine_Info> sVN_Equipment_Machine_Info { get; set; }

        public DbSet<SVN_Equipment_Status_Update_Detail> SVN_Equipment_Status_Update_Detail { get; set; }


        public DbSet<SVN_Equipment_Status_Update> SVN_Equipment_Status_Update { get; set; }

    }

}
