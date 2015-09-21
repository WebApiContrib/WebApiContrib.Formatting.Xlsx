using WebApiContrib.Formatting.Xlsx.Attributes;

namespace WebApiContrib.Formatting.Xlsx.Sample.Models
{
    [ExcelDocument("CIA World Factbook - Internet Penetration")]
    public class CiaWorldFactBookData
    {
        public string Country { get; set; }

        [ExcelColumn(Header = "Population in 2010 (estimated)", NumberFormat = "#,0.##,,\" M\"")]
        public int EstimatedPopulationIn2010 { get; set; }

        [ExcelColumn(Header = "% of world population", NumberFormat = "0%")]
        public decimal PercentOfWorldPopulation { get; set; }

        [ExcelColumn(Header = "Internet users", NumberFormat = "#,0.##,,\" M\"")]
        public int? InternetUsers { get; set; }

        [ExcelColumn(NumberFormat = "0%")]
        public decimal? Penetration { get; set; }

        public string Region { get; set; }

        [ExcelColumn(Header = "Income group")]
        public string IncomeGroup { get; set; }

        [ExcelColumn(Header = "GDP per capita", NumberFormat = "$#,###")]
        public int? GdpPerCapita { get; set; }
    }
}