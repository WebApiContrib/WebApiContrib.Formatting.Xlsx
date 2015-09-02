using System;
using System.Web;

namespace WebApiContrib.Formatting.Xlsx.Utils {
    public class HttpContextFactory
    {
        private static HttpContextBase currentContext;

        public static HttpContextBase Current
        {
            get
            {
                if (currentContext != null) return currentContext;

                if (HttpContext.Current == null)
                    throw new InvalidOperationException("HttpContext is not available.");

                return new HttpContextWrapper(HttpContext.Current);
            }
        }

        public static void SetCurrentContext(HttpContextBase context)
        {
            HttpContextFactory.currentContext = context;
        }
    }
}