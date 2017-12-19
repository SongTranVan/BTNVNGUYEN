using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectTeam_v1._0
{
    public class Khoa
    {
        [Key]
        public string MaKhoa { get; set; }
        [Required]
        public string TenKhoa { get; set; }
        public virtual ICollection<Lop> Lops { get; set; }
        public Khoa()
        {

        }

    }
}
