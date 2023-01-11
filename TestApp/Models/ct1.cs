using System.ComponentModel;
using System.Data;
using System.Reflection;

namespace TestApp.Models
{
    public static class ct1
    {
        public static DataTable ToDataTable<T>(IList<T> data, Dictionary<string, string> activeColumnList = null)
        {
            PropertyDescriptorCollection propTemps =
                TypeDescriptor.GetProperties(typeof(T));

            List<PropertyDescriptor> propsTemp = new List<PropertyDescriptor>();

            foreach (PropertyDescriptor temp in propTemps)
            {
                if (activeColumnList == null)
                    propsTemp.Add(temp);
                else if (activeColumnList.ContainsKey(temp.Name))
                    propsTemp.Add(temp);

            }

            List<PropertyDescriptor> props = new List<PropertyDescriptor>();
            if (activeColumnList != null)
            {
                foreach (var key in activeColumnList.Keys)
                {
                    var pTemp = propsTemp.FirstOrDefault(i => i.Name == key);
                    if (pTemp != null)
                        props.Add(pTemp);
                }
            }
            else
            {
                props = propsTemp;
            }

            DataTable table = new DataTable();
            try
            {
                Type Propiedad = null;
                for (int i = 0; i < props.Count; i++)
                {
                    PropertyDescriptor prop = props[i];
                    Propiedad = prop.PropertyType;
                    if (Propiedad.IsGenericType && Propiedad.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        Propiedad = Nullable.GetUnderlyingType(Propiedad);
                    }

                    table.Columns.Add(activeColumnList == null ? prop.Name : activeColumnList[prop.Name], Propiedad);
                }
                object[] values = new object[props.Count];
                foreach (T item in data)
                {
                    for (int i = 0; i < values.Length; i++)
                    {
                        values[i] = props[i].GetValue(item);
                    }
                    table.Rows.Add(values);
                }
            }
            catch (Exception ex)
            {

            }
            return table;
        }



        public static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }
        public static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    else
                        continue;
                }
            }
            return obj;
        }
    }
}
