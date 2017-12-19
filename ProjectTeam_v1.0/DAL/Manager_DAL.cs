using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using XLS = Microsoft.Office.Interop.Excel.Application;

namespace ProjectTeam_v1._0
{
    public class Manager_DAL
    {
        public Manager db;
        public Manager_DAL()
        {
            db = new Manager();
        }

        public bool Check_MSV_DAL(string MaSv)
        {
            var s = db.SinhViens.Where(p => p.MaSinhVien == MaSv).Select(p => p);
            if (s.Any())
            {
                return false;
            }
            else return true;
        }

        public void ReadExcel_SinhVien_DAL(string path)
        {
            string filename = path;
            Excel.Application application = new Excel.Application();
            application.Workbooks.Open(filename);
            foreach (Excel.Worksheet worksheet in application.Worksheets)
            {
                int i = 5;
                while (worksheet.Range["B"+i].Value != null)
                {
                    string tmpMsv = worksheet.Range["B" + i].Value;
                    if (Check_MSV_DAL(tmpMsv) == false)
                    {

                        SinhVien sv = new SinhVien();
                        sv.MaSinhVien = worksheet.Range["B" + i].Value;
                        sv.TenSinhVien = worksheet.Range["C" + i].Value;
                        sv.NgaySinh = Convert.ToDateTime(worksheet.Range["D" + i].Value.ToString());
                        if (int.Parse(worksheet.Range["E" + i].Value.ToString()) == 1) sv.GioiTinh = true;
                        if (int.Parse(worksheet.Range["E" + i].Value.ToString()) == 0) sv.GioiTinh = false;
                        sv.DanToc = worksheet.Range["F" + i].Value;
                        sv.TonGiao = worksheet.Range["G" + i].Value;
                        sv.SoCMND = worksheet.Range["H" + i].Value;
                        sv.SoDienThoai = worksheet.Range["I" + i].Value;
                        sv.QueQuan = worksheet.Range["J" + i].Value;
                        sv.DiaChiTamTru = worksheet.Range["K" + i].Value;
                        sv.NienKhoa = int.Parse(worksheet.Range["L" + i].Value.ToString());
                        sv.MaLop = worksheet.Range["M" + i].Value;
                        db.SinhViens.Add(sv);
                        db.SaveChanges();
                        i++;
                    }
                    else
                    {
                        i++;
                    }
                }
            }
        }

        public void ReadExcel_KiHoc_DAL(string path)
        {
            string filename = path;
            Excel.Application application = new Excel.Application();
            application.Workbooks.Open(filename);
            foreach (Excel.Worksheet worksheet in application.Worksheets)
            {
                int i = 2;
                while (worksheet.Range["B" + i].Value != null)
                {
                    string tmpMsv = worksheet.Range["B" + i].Value;
                    if (Check_MSV_DAL(tmpMsv) == false)
                    {

                        KiHoc kh = new KiHoc();
                        kh.MaSinhVien = worksheet.Range["B" + i].Value;
                        kh.Ki = worksheet.Range["C" + i].Value;
                        kh.DiemRenLuyen = double.Parse(worksheet.Range["D" + i].Value.ToString());
                        kh.DiemTrungBinh = double.Parse(worksheet.Range["E" + i].Value.ToString());
                        kh.CoDuocHocBong = bool.Parse(worksheet.Range["F" + i].Value.ToString());
                        db.KiHocs.Add(kh);
                        db.SaveChanges();
                        i++;
                    }
                    else
                    {
                        i++;
                    }
                }
            }
        }
        public void XuatExcel_DAL(DataGridView g, string duongDan, string tenTap)
        {
            XLS obj = new XLS();
            obj.Application.Workbooks.Add(Type.Missing);
            obj.Columns.ColumnWidth = 25;
            for (int i = 1; i < g.Columns.Count + 1; i++) { obj.Cells[1, i] = g.Columns[i - 1].HeaderText; }
            for (int i = 0; i < g.Rows.Count; i++)
            {
                for (int j = 0; j < g.Columns.Count; j++)
                {
                    if (g.Rows[i].Cells[j].Value != null) { obj.Cells[i + 2, j + 1] = g.Rows[i].Cells[j].Value.ToString(); }
                }
            }
            obj.ActiveWorkbook.SaveCopyAs(duongDan + tenTap + ".xlsx");
            obj.ActiveWorkbook.Saved = true;
        }
    }
}
