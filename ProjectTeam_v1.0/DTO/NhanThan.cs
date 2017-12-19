using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectTeam_v1._0
{
     public class NhanThan
    {
        // Phần nhân thân ở trên trang đào tạo thì Họ tên, Quan hệ, Địa chỉ là một chuỗi string
        // Nếu muốn đơn giản hóa(để làm cho kịp) thì có thể gộp các thuộc tính dưới lại thành 1 thuộc tính duy nhất
        // Mỗi sinh viên có một nhân thân
        [Key]
        public string MaNhanThan { get; set; }
        //[Required]
        //public string HoTen { get; set; }
        //[Required]
        //public string QuanHe { get; set; }
        //[Required]
        //public string DiaChi { get; set; }
        //[Required]
        //public string SoDienThoai { get; set; }
        [Required]
        public string MaSinhVien { get; set; }

        [ForeignKey("MaSinhVien")]
        public virtual SinhVien SinhVien { get; set; }

    }
}
