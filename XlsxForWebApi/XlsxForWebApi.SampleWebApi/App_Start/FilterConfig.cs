using System.Web;
using System.Web.Mvc;

namespace XlsxForWebApi.SampleWebApi
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
