using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace MachineStatusUpdate.Models
{
    public class SVN_Equipment_Machine_Test
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        public string Project { get; set; }

        public string SVNCode { get; set; }
    }
}