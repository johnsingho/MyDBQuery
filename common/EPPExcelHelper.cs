using OfficeOpenXml;
using System;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;

namespace Common
{
    /// <summary>
    /// EPPExcelHelper
    /// use EPPlus
    /// By H.Z.XIN
    /// Modified:
    ///     2018-08-09 add DataReader, 分批导入数据库
    /// 
    /// </summary>
    public class EPPExcelHelper
    {
        //使用DataReader来导出
        // 1048576行是默认的极限
        public static byte[] BuilderExcel(DataTable table, string dateFormat = "yyyy-MM-dd HH:mm:ss")
        {
            OfficeOpenXml.ExcelPackage pkg = new ExcelPackage();
            var sheet1 = pkg.Workbook.Worksheets.Add("sheet1");
            if (table == null)
            {
                return null;
            }
            sheet1.Cells[1, 1].LoadFromDataTable(table, true);
            var len = table.Columns.Count;
            for (int i = 0; i < len; i++)
            {
                if (table.Columns[i].DataType == typeof(DateTime))
                {
                    sheet1.Column(i + 1).Style.Numberformat.Format = dateFormat;
                }
            }
            //int colCount = table.Columns.Count;
            //int rowCount = table.Rows.Count;
            //for (int i = 0; i < colCount; i++) {
            //    sheet1.Cells[1, i + 1].Value = table.Columns[i].ColumnName;
            //}
            //if (table.Rows.Count > 0) {
            //    for (int i = 0; i < rowCount; i++) {
            //        for (int j = 0; j < colCount; j++) {
            //            sheet1.Cells[2 + i, j + 1].Value = string.Format("{0}", table.Rows[i][j]);
            //        }
            //    }
            //}
            return pkg.GetAsByteArray();
        }

        //使用DataReader来导出
        // 1048576行是默认的极限
        public static bool BuilderExcel(string sTarFile, DbDataReader dtReader, string dateFormat = "yyyy-MM-dd HH:mm:ss")
        {
            var fi = new FileInfo(sTarFile);
            try
            {
                using (var pkg = new ExcelPackage(fi))
                {
                    var sheet1 = pkg.Workbook.Worksheets.Add("sheet1");
                    if (dtReader == null)
                    {
                        return false;
                    }
                    sheet1.Cells[1, 1].LoadFromDataReader(dtReader, true); 
                    var len = dtReader.FieldCount;
                    for (int i = 0; i < len; i++)
                    {
                        if (dtReader.GetFieldType(i) == typeof(DateTime))
                        {
                            sheet1.Column(i + 1).Style.Numberformat.Format = dateFormat;
                        }
                    }

                    pkg.Save();
                    return true;
                }
            }
            catch(Exception ex)
            {
                //Row out of range  -- ExcelPackage.MaxRows=1048576
                /* 唯有使用分页
                select* from (
                 SELECT
                    ROW_NUMBER() OVER(ORDER BY[ID]) AS rowNumber
                      ,[ID]
                      ,[dates]
                      ,[detailtime]
                      ,[tem]
                      ,[hum]
                      ,[temperaturelow]
                      ,[temperaturehigh]
                      ,[Humiditylow]
                      ,[Humidityhigh]
                      ,[machineno]
                      ,[assettype]
                      ,[uploadtime]
                      ,[actions]
                      ,[strowner]
                      ,[status]
                      ,[closedate]
                FROM[erecordcontrol].[dbo].[Icetemperaturedata]
                where machineno between 1 and 6
                and('2017-01-01' <= uploadtime and uploadtime <= GETDATE())
                ) as tq
                where tq.rowNumber between 0 and 1000000
                */

                return false;
            }
        }


        public static DataTable ReadExcel(System.IO.Stream stream)
        {
            using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(stream))
            {
                var sheet = package.Workbook.Worksheets[1];
                int colStart = sheet.Dimension.Start.Column;    //工作区开始列
                int colEnd = sheet.Dimension.End.Column;        //工作区结束列
                int rowStart = sheet.Dimension.Start.Row;       //工作区开始行号
                int rowEnd = sheet.Dimension.End.Row;           //工作区结束行号
                DataTable dt = new DataTable();
                for (int i = colStart; i <= colEnd; i++)
                {
                    var val = string.Format("{0}", sheet.Cells[rowStart, i].Value).Trim();
                    if (val == "") { val = "C" + i; }
                    dt.Columns.Add(val, typeof(string));
                }
                int s = colStart;
                for (int i = rowStart + 1; i <= rowEnd; i++)
                {
                    DataRow rw = dt.NewRow();
                    for (int j = colStart; j <= colEnd; j++)
                    {
                        var val = sheet.Cells[i, j].Value;
                        var fmtID = sheet.Cells[i, j].Style.Numberformat.NumFmtID;//
                        var fmt = sheet.Cells[i, j].Style.Numberformat.Format;
                        if (null == val)
                        {
                            rw[j - s] = null;
                        }
                        else
                        {
                            if (IsDatetimeFmt(fmtID, fmt))
                            {
                                var bCast = false;
                                DateTime tim = default(DateTime);
                                if (!string.IsNullOrEmpty(val.ToString()))
                                {
                                    tim = sheet.Cells[i, j].GetValue<DateTime>();
                                    bCast = true;
                                }

                                if (bCast)
                                {
                                    val = string.Format("{0:yyyy-MM-dd hh:mm:ss}", tim);
                                }
                                else
                                {
                                    rw[j - s] = null;
                                    continue;//!
                                }
                            }
                            rw[j - s] = string.Format("{0}", val).Trim();
                        }
                    }
                    dt.Rows.Add(rw);
                }
                return dt;
            }
        }

        private static readonly int[] DateFmtID = new int[] {14,15,16,17,22 };
        private static readonly string[] DATETIME_FMT = new string[] {
                            "mm-dd-yy",
                            "mm/dd/yyyy",
                            "dd-mmm-yy",
                            "mm/dd/yyyy hh:mm",
                            "mm/dd/yyyy hh:mm:ss",
                            "dd-mmm-yy",
                            "m/d/yy\\ h:mm;@",
                            "m/d/yyyy h:mm",
                            "m/d/yyyy H:mm",
                            "m/d/yyyy H:mm:ss",
                            "yyyy\\-mm\\-dd",
                            "yyyy\\-mm\\-dd hh:mm",
                            "yyyy\\-mm\\-dd hh:mm:ss",
                            "yyyy\\-MM\\-dd HH:mm",
                            "yyyy\\-MM\\-dd HH:mm:ss",
                            "MM/dd/yyyy hh:mm:ss AM/PM"
            };
        private static bool IsDatetimeFmt(int fmtID, string strFmt)
        {
            if(0==string.Compare("General", strFmt, true))
            {
                return false;
            }
            if (DateFmtID.Contains(fmtID))
            {
                return true;
            }
            if (DATETIME_FMT.Contains(strFmt))
            {
                return true;
            }
            if (fmtID == 173
                && strFmt.IndexOf("m", StringComparison.InvariantCultureIgnoreCase) > -1
                && strFmt.IndexOf("y", StringComparison.InvariantCultureIgnoreCase) > -1
                && strFmt.IndexOf("d", StringComparison.InvariantCultureIgnoreCase) > -1
                )
            {
                return true;
            }
            
            return false;
        }

        public object GetMegerValue(ExcelWorksheet wSheet, int row, int column)
        {
            string range = wSheet.MergedCells[row, column];
            if (range == null)
            {
                return wSheet.Cells[row, column].Value;
            }
            else
            {
                return wSheet.Cells[(new ExcelAddress(range)).Start.Row, (new ExcelAddress(range)).Start.Column].Value;
            }
        }

        //从Excel读取，然后分批地写入数据库
        public delegate bool FBatchDataHandle(DataTable dt, ref long nHandles, out string sErr);
        public static long ReadExcel(System.IO.FileInfo file, string sTarTable,
                                    FBatchDataHandle handler,  out string sErr, int nBatchSize=4096)
        {
            long nHandles = 0;
            sErr = string.Empty;
            if (null==file || null==handler)
            {
                return nHandles;
            }

            using (var package = new OfficeOpenXml.ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets[1];
                int colStart = sheet.Dimension.Start.Column;    //工作区开始列
                int colEnd = sheet.Dimension.End.Column;        //工作区结束列
                int rowStart = sheet.Dimension.Start.Row;       //工作区开始行号
                int rowEnd = sheet.Dimension.End.Row;           //工作区结束行号
                DataTable dt = new DataTable();
                dt.TableName = sTarTable;
                for (int i = colStart; i <= colEnd; i++)
                {
                    var val = string.Format("{0}", sheet.Cells[rowStart, i].Value).Trim();
                    if (val == "") { val = "C" + i; }
                    dt.Columns.Add(val, typeof(string));
                }

                bool bCont = true;
                int s = colStart;
                for (int i = rowStart + 1; i <= rowEnd; i++)
                {
                    DataRow rw = dt.NewRow();
                    for (int j = colStart; j <= colEnd; j++)
                    {
                        var val = sheet.Cells[i, j].Value;
                        var fmtID = sheet.Cells[i, j].Style.Numberformat.NumFmtID;//
                        var fmt = sheet.Cells[i, j].Style.Numberformat.Format;
                        if (null == val)
                        {
                            rw[j - s] = null;
                        }
                        else
                        {
                            if (IsDatetimeFmt(fmtID, fmt))
                            {
                                var bCast = false;
                                DateTime tim=default(DateTime);
                                if (!string.IsNullOrEmpty(val.ToString()))
                                {
                                    tim = sheet.Cells[i, j].GetValue<DateTime>();
                                    bCast = true;
                                }

                                if (bCast)
                                {
                                    val = string.Format("{0:yyyy-MM-dd hh:mm:ss}", tim);
                                }
                                else
                                {
                                    rw[j-s] = null;
                                    continue;//!
                                }
                            }
                            rw[j - s] = string.Format("{0}", val).Trim();
                        }
                    }
                    dt.Rows.Add(rw);
                    if(dt.Rows.Count >= nBatchSize)
                    {
                        bCont = handler(dt, ref nHandles, out sErr);//handle
                        if (!bCont)
                        {
                            break; //有错发生
                        }
                        dt.Rows.Clear();
                    }
                }

                if (dt.Rows.Count>=0 && bCont)
                {
                    handler(dt, ref nHandles, out sErr);//handle
                    dt.Rows.Clear();
                }
                return nHandles;
            }
        }

        public static DataTable ReadExcel(System.IO.FileInfo file)
        {
            using (System.IO.Stream stream = file.OpenRead())
            {
                return ReadExcel(stream);
            }
        }
    }
}
