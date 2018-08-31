using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections;
using System.Reflection;

namespace ComprehensiveHardwareInventory
{
    public class Tools
    {
        #region 打开保存excel对话框返回文件名
        public static string SaveExcelFileDialog()
        {
            var sfd = new Microsoft.Win32.SaveFileDialog()
            {
                DefaultExt = ".xlsx",
                Filter = "excel files(*.xlsx)|*.xls|All files(*.*)|*.*",
                FilterIndex = 1
            };

            if (sfd.ShowDialog() != true)
                return null;
            return sfd.FileName;
        }
        #endregion
        #region 打开excel对话框返回文件名
        public static string OpenExcelFileDialog()
        {
            var ofd = new Microsoft.Win32.OpenFileDialog()
            {
                DefaultExt = "xls",
                Filter = "excel files(*.xls)|*.xls|All files(*.*)|*.*",
                FilterIndex = 1
            };

            if (ofd.ShowDialog() != true)
                return null;
            return ofd.FileName;
        }
        #endregion
        #region 读excel
        public static DataTable ImportExcelFile()
        {
            DataTable dt = new DataTable();

            //打开excel对话框
            var filepath = OpenExcelFileDialog();
            if (filepath != null)
            {

                HSSFWorkbook hssfworkbook = null;
                #region//初始化信息
                try
                {
                    using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read))
                    {
                        hssfworkbook = new HSSFWorkbook(file);
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
                #endregion

                var sheet = hssfworkbook.GetSheetAt(0);
                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

                for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
                {
                    dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());
                }
                while (rows.MoveNext())
                {
                    HSSFRow row = (HSSFRow)rows.Current;
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        var cell = row.GetCell(i);
                        if (cell == null)
                        {
                            dr[i] = "";
                        }
                        else
                        {
                            if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            {
                                if (HSSFDateUtil.IsCellDateFormatted(cell))
                                {
                                    dr[i] = cell.DateCellValue;
                                }
                                else
                                {
                                    dr[i] = cell.NumericCellValue;
                                }
                            }
                            else if (cell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                            {
                                dr[i] = cell.BooleanCellValue;
                            }
                            else
                            {
                                dr[i] = cell.StringCellValue;
                            }
                        }
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
        #endregion



        #region list转datatable
        public static DataTable ListToDataTable<T>(IEnumerable<T> c)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray());
            if (c.Count() > 0)
            {
                for (int i = 0; i < c.Count(); i++)
                {
                    ArrayList tempList = new ArrayList();
                    foreach (PropertyInfo item in props)
                    {
                        object obj = item.GetValue(c.ElementAt(i), null);
                        tempList.Add(obj);
                    }
                    dt.LoadDataRow(tempList.ToArray(), true);
                }
            }
            return dt;
        }
        #endregion
        #region 写入excel
        public static bool WriteExcel<T>(IList<T> list)
        {

            //打开保存excel对话框
            var filepath = SaveExcelFileDialog();
            if (filepath == null)
                return false;

            var dt = ListToDataTable<T>(list);

            if (!string.IsNullOrEmpty(filepath) && null != dt && dt.Rows.Count > 0)
            {
                NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
                NPOI.SS.UserModel.ISheet sheet = book.CreateSheet("Sheet1");

                NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    NPOI.SS.UserModel.IRow row2 = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        row2.CreateCell(j).SetCellValue(Convert.ToString(dt.Rows[i][j]));
                    }
                }
                // 写入到客户端  
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    book.Write(ms);
                    using (FileStream fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                    {
                        byte[] data = ms.ToArray();
                        fs.Write(data, 0, data.Length);
                        fs.Flush();
                    }
                    book = null;
                }
            }
            return true;
        }
        #endregion
    }
}
