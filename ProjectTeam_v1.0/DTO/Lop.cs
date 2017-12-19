using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectTeam_v1._0
{
    public class Lop
    {
        [Key]
        public string MaLop { get; set; }
        [Required]
        public string TenLop { get; set; }
        [Required]
        public string MaKhoa { get; set; }
        //[Required]
        //public string TenGiaoVien { get; set; }
        [ForeignKey("MaKhoa")]
        public virtual Khoa Khoa { get; set; }
        public virtual ICollection<SinhVien> SinhViens { get; set; }

        public Lop()
        {

        }

    }
}
