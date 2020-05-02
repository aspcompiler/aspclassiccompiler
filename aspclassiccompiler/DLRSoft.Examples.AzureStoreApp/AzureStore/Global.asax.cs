// ----------------------------------------------------------------------------------
// Microsoft Developer & Platform Evangelism
// 
// Copyright (c) Microsoft Corporation. All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
// EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES 
// OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
// ----------------------------------------------------------------------------------
// The example companies, organizations, products, domain names,
// e-mail addresses, logos, people, places, and events depicted
// herein are fictitious.  No association with any real company,
// organization, product, domain name, email address, logo, person,
// places, or events is intended or should be inferred.
// ----------------------------------------------------------------------------------

namespace MVCAzureStore
{
    using System;
    using System.Collections.Generic;
    using System.Web.Mvc;
    using System.Web.Routing;
    using System.Web.Security;
    using Dlrsoft.Asp.Mvc;

    // Note: For instructions on enabling IIS6 or IIS7 classic mode, 
    // visit http://go.microsoft.com/?LinkId=9394801

    public class MvcApplication : System.Web.HttpApplication
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute(
                "Default",                                              // Route name
                "{controller}/{action}/{id}",                           // URL with parameters
                new { controller = "Home", action = "Index", id = "" }  // Parameter defaults
            );

        }

        protected void Application_Start()
        {
            RegisterRoutes(RouteTable.Routes);
            ViewEngines.Engines.Add(new AspViewEngine());
        }

        private static bool initialized = false;
        private static object gate = new object();

        // The Windows Azure fabric runs IIS 7.0 in integrated mode. In integrated 
        // mode, the Application_Start event does not support access to the request 
        // context or to the members of the RoleManager class provided by the Windows 
        // Azure SDK runtime API. If you are writing an ASP.NET application that 
        // accesses the request context or calls methods of the RoleManager class from 
        // the Application_Start event, you should modify it to initialize in the 
        // Application_BeginRequest event instead.
        protected void Application_BeginRequest(object sender, EventArgs e)
        {
            if (initialized)
            {
                return;
            }

            lock (gate)
            {
                if (initialized)
                {
                    return;
                }

                LoadProducts();

                // Initialize application roles here 
                // (requires access to RoleManager to read configuration)
                if (!Roles.RoleExists("Home"))
                {
                    Roles.CreateRole("Home");
                }

                if (!Roles.RoleExists("Enterprise"))
                {
                    Roles.CreateRole("Enterprise");
                }

                initialized = true;
            }
        }

        private void LoadProducts()
        {
            this.Application["Products"] = new List<string> {
                                        "Microsoft Office 2007 Ultimate",
                                        "Microsoft Office Communications Server Enterprise CAL",
                                        "Microsoft Core CAL - License & software assurance - 1 user CAL",
                                        "Windows Server 2008 Enterprise",
                                        "Windows Vista Home Premium (Upgrade)",
                                        "Windows XP Home Edition w/SP2 (OEM)",
                                        "Windows Home Server - 10 Client (OEM License)",
                                        "Console XBOX 360 Arcade" 
                                    };
        }
    }
}