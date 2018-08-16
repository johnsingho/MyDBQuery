using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace MyDBQuery.common
{
	/// <summary>
    /// DateTimeHelper
    /// By H.Z.XIN
    /// Modified:
    ///     2018-08-13 整理
    /// 
    /// </summary>
    public static class DataTableHelper
    {
        public static bool IsEmptyDataTable(DataTable dt)
        {
            return (null == dt || 0 == dt.Rows.Count);
        }
        public static bool IsEmptyDataSet(DataSet ds)
        {
            if (null == ds || 0 == ds.Tables.Count)
            {
                return true;
            }
            return IsEmptyDataTable(ds.Tables[0]);
        }
		public static DataTable GetDataTable0(DataSet ds)
        {
            var dt = ds.Tables[0];
            return dt;
        }
        public static DataRow GetDataSet_Row0(DataSet ds)
        {
            var dt = ds.Tables[0];
            return dt.Rows[0];
        }

		
        public static Hashtable DataTableToHashtableByKeyValue(DataTable dt, string keyField, string valFiled)
        {
            Hashtable ht = new Hashtable();
            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string key = dr[keyField].ToString();
                    ht[key] = dr[valFiled];
                }
            }
            return ht;
        }
        
        public static IList<Hashtable> DataTableToArrayList(DataTable dt)
        {
            IList<Hashtable> result;
            if (dt == null)
            {
                result = new List<Hashtable>();
            }
            else
            {
                IList<Hashtable> datas = new List<Hashtable>();
                foreach (DataRow dr in dt.Rows)
                {
                    Hashtable ht = DataTableHelper.DataRowToHashTable(dr);
                    datas.Add(ht);
                }
                result = datas;
            }
            return result;
        }

        public static Hashtable DataTableToHashtable(DataTable dt)
        {
            Hashtable ht = new Hashtable();
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    string key = dt.Columns[i].ColumnName;
                    ht[key.ToUpper()] = dr[key];
                }
            }
            return ht;
        }

        public static Hashtable DataRowToHashTable(DataRow dr)
        {
            Hashtable htReturn = new Hashtable(dr.ItemArray.Length);
            foreach (DataColumn dc in dr.Table.Columns)
            {
                htReturn.Add(dc.ColumnName, dr[dc.ColumnName]);
            }
            return htReturn;
        }
		
        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }   
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }
        
        public static IList DataTableToIList<T>(DataTable dt)
        {
            IList list = new List<T>();
            foreach (DataRow dr in dt.Rows)
            {
                T obj = Activator.CreateInstance<T>();
                PropertyInfo[] propertys = obj.GetType().GetProperties();
                PropertyInfo[] array = propertys;
                int i = 0;
                while (i < array.Length)
                {
                    PropertyInfo pi = array[i];
                    string tempName = pi.Name;
                    if (dt.Columns.Contains(tempName))
                    {
                        if (pi.CanWrite)
                        {
                            object value = dr[tempName];
                            if (value != DBNull.Value)
                            {
                                pi.SetValue(obj, value, null);
                            }
                        }
                    }
                IL_B2:
                    i++;
                    continue;
                    goto IL_B2;
                }
                list.Add(obj);
            }
            return list;
        }        
		
		public static DataTable GetPagedTable(DataTable dt, int PageIndex, int PageSize)
        {
            DataTable result;
            if (PageIndex == 0)
            {
                result = dt;
            }
            else
            {
                DataTable newdt = dt.Copy();
                newdt.Clear();
                int rowbegin = (PageIndex - 1) * PageSize;
                int rowend = PageIndex * PageSize;
                if (rowbegin >= dt.Rows.Count)
                {
                    result = newdt;
                }
                else
                {
                    if (rowend > dt.Rows.Count)
                    {
                        rowend = dt.Rows.Count;
                    }
                    for (int i = rowbegin; i <= rowend - 1; i++)
                    {
                        DataRow newdr = newdt.NewRow();
                        DataRow dr = dt.Rows[i];
                        foreach (DataColumn column in dt.Columns)
                        {
                            newdr[column.ColumnName] = dr[column.ColumnName];
                        }
                        newdt.Rows.Add(newdr);
                    }
                    result = newdt;
                }
            }
            return result;
        }
		
        /// <summary>
        ///将DataTable转换为标准的CSV
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <returns>返回标准的CSV</returns>
        public static string DataTableToCsv(DataTable dt)
        {
            //以半角逗号（即,）作分隔符，列为空也要表达其存在。
            //列内容如存在半角逗号（即,）则用半角引号（即""）将该字段值包含起来。
            //列内容如存在半角引号（即"）则应替换成半角双引号（""）转义，并用半角引号（即""）将该字段值包含起来。
            StringBuilder sb = new StringBuilder();
            foreach (DataRow row in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    DataColumn colum = dt.Columns[i];
                    if (i != 0) { sb.Append(","); }
                    if (colum.DataType == typeof(string) && row[colum].ToString().Contains(","))
                    {
                        sb.Append("\"" + row[colum].ToString().Replace("\"", "\"\"") + "\"");
                    }
                    else sb.Append(row[colum].ToString());
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        public static List<string> GetDataTableColNames(DataTable dt)
        {
            var lst = new List<string>();
            foreach(DataColumn col in dt.Columns)
            {
                lst.Add(col.ColumnName);
            }
            return lst;
        }

    }
}
