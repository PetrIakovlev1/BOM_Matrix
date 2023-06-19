using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;

namespace MSOL_Matrix
{
    public static class ConvertFunctions
    {
        public static int GetConvertedToIntValue(object oInputValue)
        {
            int iOutValue = 0;

            if (oInputValue != DBNull.Value && oInputValue != null)
            {
                string sConvertedValue = oInputValue.ToString();

                if (!string.IsNullOrWhiteSpace(sConvertedValue))
                    int.TryParse(sConvertedValue, out iOutValue);
            }

            return iOutValue;
        }

        public static bool GetConvertedToBoolValue(object oInputValue)
        {
            bool bOutValue = false;

            if (oInputValue != DBNull.Value && oInputValue != null)
            {
                string sConvertedValue = oInputValue.ToString();

                if (!string.IsNullOrWhiteSpace(sConvertedValue))
                    if (sConvertedValue == "1")
                        bOutValue = true;
                    else
                        bool.TryParse(sConvertedValue, out bOutValue);
            }

            return bOutValue;
        }

        public static DateTime? GetConvertedToDateTimeValue(object oInputValue)
        {
            DateTime? iOutValue = null;

            if (oInputValue != DBNull.Value && oInputValue != null)
            {
                string sConvertedValue = oInputValue.ToString();
                DateTime TmpiOutValue;

                if (!string.IsNullOrWhiteSpace(sConvertedValue))
                    if (DateTime.TryParse(sConvertedValue, out TmpiOutValue))
                        iOutValue = TmpiOutValue;
            }

            return iOutValue;
        }

        public static string GetConvertedToString(object oInputValue)
        {
            string sOutValue = null;

            if (oInputValue != DBNull.Value && oInputValue != null)
                sOutValue = oInputValue.ToString();

            if (string.IsNullOrEmpty(sOutValue))
                sOutValue = null;

            return sOutValue;
        }

        public static double GetConvertedToDblValue(object oInputValue)
        {
            double dOutValue = 0;

            var culture = System.Globalization.CultureInfo.CurrentCulture;


            if (oInputValue != DBNull.Value && oInputValue != null)
            {
                string sConvertedValue = oInputValue.ToString();
                if (culture.Name == "ru-RU")
                    sConvertedValue = oInputValue.ToString().Replace(".", ",");

                if (!string.IsNullOrWhiteSpace(sConvertedValue))
                    double.TryParse(sConvertedValue, out dOutValue);
            }

            return dOutValue;
        }

        public static List<object> DtToObjList(DataTable dt)
        {
            DataTable dtTmp = dt.Copy();
            //dtTmp.Columns.RemoveAt(dtTmp.Columns.IndexOf("Date"));

            var rows = dtTmp.Rows;
            int rowCount = rows.Count;
            int colCount = dtTmp.Columns.Count;

            List<object> objs = new List<object>();
            objs.Add((from dc in dtTmp.Columns.Cast<DataColumn>()
                      select dc.ColumnName).ToArray());

            for (int i = 0; i < rowCount; i++)
            {
                objs.Add(rows[i].ItemArray);
            }

            return objs;
        }

        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();

            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

    }
}