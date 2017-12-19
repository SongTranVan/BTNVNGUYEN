using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectTeam_v1._0
{
    public class Manager_BLL
    {
        public Manager_DAL DAL;
        public Manager_BLL()
        {
            DAL = new Manager_DAL();
        }

        public void ReadExcel_SinhVien_BLL(string path)
        {
            DAL.ReadExcel_SinhVien_DAL(path);
        }

        public void ReadExcel_KiHoc_BLL(string path)
        {
            DAL.ReadExcel_KiHoc_DAL(path);
        }

        public void XuatExcel_BLL (DataGridView g, string duongDan, string tenTap)
        {
            DAL.XuatExcel_DAL(g, duongDan, tenTap);
        }
    }
}
