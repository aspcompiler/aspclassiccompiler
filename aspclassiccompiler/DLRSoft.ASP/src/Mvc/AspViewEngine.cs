using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Mvc;

namespace Dlrsoft.Asp.Mvc
{
    public class AspViewEngine : VirtualPathProviderViewEngine
    {
        public AspViewEngine()
        {
            // This is where we tell MVC where to look for our files. This says
            // to look for a file at "Views/Controller/Action.html"
            base.ViewLocationFormats = new string[] { 
                "~/Views/{1}/{0}.asp",
                "~/Views/Shared/{0}.asp"
            };

            base.PartialViewLocationFormats = new string[] { 
                "~/Views/{1}/{0}.asc",
                "~/Views/{1}/{0}.asp",
                "~/Views/Shared/{0}.asc",
                "~/Views/Shared/{0}.asp"
            };
        }

        protected override IView CreateView(ControllerContext context, string viewPath, string masterPath)
        {
            return new AspView(viewPath, masterPath);
        }

        protected override IView CreatePartialView(ControllerContext context, string partialPath)
        {
            return new AspView(partialPath, "");
        }
    }
}
