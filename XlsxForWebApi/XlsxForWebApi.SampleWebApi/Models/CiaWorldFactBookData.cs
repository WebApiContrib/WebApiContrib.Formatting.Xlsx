using XlsxForWebApi;

namespace XlsxForWebApi.SampleWebApi.Models
{
    public class CiaWorldFactBookData
    {
        public string Country { get; set; }

        [Excel(Header = "Population in 2010 (estimated)", NumberFormat = "#,0.##,,\" M\"")]
        public int EstimatedPopulationIn2010 { get; set; }

        [Excel(Header = "% of world population", NumberFormat = "0%")]
        public decimal PercentOfWorldPopulation { get; set; }

        [Excel(Header = "Internet users", NumberFormat = "#,0.##,,\" M\"")]
        public int? InternetUsers { get; set; }

        [Excel(NumberFormat = "0%")]
        public decimal? Penetration { get; set; }

        public string Region { get; set; }

        [Excel(Header = "Income group")]
        public string IncomeGroup { get; set; }

        [Excel(Header = "GDP per capita", NumberFormat = "$#,###")]
        public int? GdpPerCapita { get; set; }
    }
}