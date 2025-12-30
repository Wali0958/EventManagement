using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Web;

namespace EventManagement.Models
{
    public class LoginModels
    {

        [Required(ErrorMessage = "email is required.")]
        [MaxLength(70)]
        public string UserName { get; set; }
        [Required(ErrorMessage = "password is required.")]
        [MaxLength(15)]
        public string Pass { get; set; }
    }
    public static class CommonFunctions
    {
        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                dataTable.Columns.Add(prop.Name, type);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
        public static List<T> ToList<T>(this DataTable table) where T : new()
        {
            IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
            List<T> result = new List<T>();

            foreach (var row in table.Rows)
            {
                var item = CreateItemFromRow<T>((DataRow)row, properties);
                result.Add(item);
            }

            return result;
        }
        private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties) where T : new()
        {
            T item = new T();
            try
            {
                foreach (var property in properties)
                {
                    if (property.PropertyType == typeof(System.DayOfWeek))
                    {
                        DayOfWeek day = (DayOfWeek)Enum.Parse(typeof(DayOfWeek), row[property.Name].ToString());
                        property.SetValue(item, day, null);
                    }
                    else
                    {
                        if (row[property.Name] == DBNull.Value)
                            property.SetValue(item, null, null);
                        else
                            property.SetValue(item, row[property.Name], null);
                    }
                }

                return item;
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
                return item;
            }
        }
        public static T ToBindData<T>(DataTable dt)
        {
            // Create object
            var ob = Activator.CreateInstance<T>();
            try
            {
                DataRow dr = dt.Rows[0];

                // Get all columns' name
                List<string> columns = new List<string>();
                foreach (DataColumn dc in dt.Columns)
                {
                    columns.Add(dc.ColumnName);
                }

                // Get all fields
                var fields = typeof(T).GetFields();
                foreach (var fieldInfo in fields)
                {
                    try
                    {
                        if (columns.Contains(fieldInfo.Name))
                        {
                            // Fill the data into the field
                            if (dr[fieldInfo.Name] != null && dr[fieldInfo.Name].ToString() != "")
                            {
                                fieldInfo.SetValue(ob, dr[fieldInfo.Name]);
                            }
                        }
                    }
                    catch (Exception ee)
                    {
                        ee.Message.ToString();
                    }
                }

                // Get all properties
                var properties = typeof(T).GetProperties();
                foreach (var propertyInfo in properties)
                {
                    try
                    {
                        if (propertyInfo.Name == "WhenOn")
                        {

                        }

                        if (columns.Contains(propertyInfo.Name))
                        {
                            // Fill the data into the property   
                            if (dr[propertyInfo.Name] != null && dr[propertyInfo.Name].ToString() != "")
                            {
                                propertyInfo.SetValue(ob, dr[propertyInfo.Name]);
                            }
                        }
                    }
                    catch (Exception ee2)
                    {
                        ee2.Message.ToString();
                    }
                }

            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
            return ob;
        }
    }
}