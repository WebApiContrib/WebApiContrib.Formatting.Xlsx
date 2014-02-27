using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Http;
using WebApiContrib.Formatting.Xlsx.Sample.Models;

namespace WebApiContrib.Formatting.Xlsx.Sample.Controllers
{
    public class ValuesController : ApiController
    {
        /// <summary>
        /// Get a list of countries with data from the CIA World Factbook.
        /// </summary>
        public IEnumerable<CiaWorldFactBookData> Get()
        {
            var path = AppDomain.CurrentDomain.BaseDirectory + @"\cia-data.txt";
            var data = File.ReadAllLines(path, Encoding.UTF8);

            var ciaData = from line in data
                          select line.Split('|')
                              into row
                              select new CiaWorldFactBookData()
                              {
                                  Country = row[0],
                                  EstimatedPopulationIn2010 = int.Parse(row[1]),
                                  PercentOfWorldPopulation = decimal.Parse(row[2]),
                                  InternetUsers = row[3].ToNullable<int>(),
                                  Penetration = row[4].ToNullable<decimal>(),
                                  Region = row[5],
                                  IncomeGroup = row[6],
                                  GdpPerCapita = row[7].ToNullable<int>()
                              };

            return ciaData;
        }
    }

    internal static class StringExtensions
    {
        public static Nullable<T> ToNullable<T>(this string s) where T : struct
        {
            Nullable<T> result = new Nullable<T>();
            try
            {
                if (!string.IsNullOrEmpty(s) && s.Trim().Length > 0)
                {
                    TypeConverter conv = TypeDescriptor.GetConverter(typeof(T));
                    result = (T)conv.ConvertFrom(s);
                }
            }
            catch { }
            return result;
        }
    }
}