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
    public static class DataTableHelper
    {
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
