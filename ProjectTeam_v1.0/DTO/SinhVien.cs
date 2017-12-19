using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectTeam_v1._0
{
    public class SinhVien
    {
        // Phần này ta làm các thuộc tính theo file excel của ta gửi Sinh, nếu dài dòng quá thì bỏ bớt
        [Key]
        public string MaSinhVien { get; set; }
        [Required]
        public string MaLop { get; set; }
        [Required]
        public string TenSinhVien { get; set; }
        [Required]
        public bool GioiTinh { get; set; }
        [Required]
        public string DanToc { get; set; }
        [Required]
        public string TonGiao { get; set; }
        [Required]
        public string SoCMND { get; set; }
        [Required]
        public string SoDienThoai { get; set; }   
        [Required]
        public string QueQuan { get; set; }
        [Required]
        public string DiaChiTamTru { get; set; }
        [Required]
        public DateTime NgaySinh { get; set; }
        [Required]
        public int NienKhoa { get; set; }
        [ForeignKey("MaLop")]
        public virtual Lop Lop { get; set; }
        public virtual ICollection<KiHoc> KiHoc { get; set; }



        public SinhVien()
        {

        }

    }
}
