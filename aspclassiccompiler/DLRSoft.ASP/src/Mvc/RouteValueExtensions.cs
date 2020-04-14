using System;
using System.Collections.Generic;
using System.Web.Mvc;
using System.Web.Routing;
using System.Text;

namespace Dlrsoft.Asp.Mvc
{
    public static class RouteValueExtensions
    {
        public static RouteValueDictionary RouteValue(this HtmlHelper htmlHelper, string key, object value)
        {
            RouteValueDictionary routeValue = new RouteValueDictionary();
            routeValue.Add(key, value);
            return routeValue;
        }

        public static RouteValueDictionary RouteValue(this HtmlHelper htmlHelper, string key1, object value1, string key2, object value2)
        {
            RouteValueDictionary routeValue = new RouteValueDictionary();
            routeValue.Add(key1, value1);
            routeValue.Add(key2, value2);
            return routeValue;
        }

    }
}
