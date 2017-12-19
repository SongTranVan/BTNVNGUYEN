using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectTeam_v1._0
{
    public class KiHoc
    {
        // Kì học là thuộc tính đa trị, cần có 2 khóa chính
        // Khóa 1 là Kì chạy từ 1-14
        // Khóa 2 là Mã sinh viên, cũng là khóa ngoại đến SinhVien
        [Key]
        [Column(Order = 0)]
        public string Ki { get; set; }
        [Key]
        [Column(Order = 1)]
        public string MaSinhVien { get; set; }
        [Required]
        public int DiemRenLuyen { get; set; }
        [Required]
        public double DiemTrungBinh { get; set; }
        [Required]
        public bool CoDuocHocBong { get; set; }
        [ForeignKey("MaSinhVien")]
        public virtual SinhVien SinhVien { get; set; }


    }
}
