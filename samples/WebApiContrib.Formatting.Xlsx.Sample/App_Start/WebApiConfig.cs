using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web.Http;

namespace WebApiContrib.Formatting.Xlsx.Sample
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services

            config.Formatters.Clear();
            var formatter = new XlsxMediaTypeFormatter(
                autoFilter: true,
                freezeHeader: true,
                headerHeight: 25f,
                cellStyle: (ExcelStyle s) => {
                    s.Font.SetFromFont(new Font("Segoe UI", 13f, FontStyle.Regular));
                },
                headerStyle: (ExcelStyle s) => {
                    s.Fill.PatternType = ExcelFillStyle.Solid;
                    s.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 114, 51));
                    s.Font.Color.SetColor(Color.White);
                    s.Font.Size = 15f;
                }
            );

            config.Formatters.Add(formatter);

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
