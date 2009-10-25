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

namespace MVCAzureStore.Controllers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Web.Mvc;

    [HandleError]
    [Authorize]
    public class HomeController : Controller
    {
        public ActionResult About()
        {
            return View();
        }

        public ActionResult Index()
        {
            var products = this.HttpContext.Application["Products"] as List<string>;
            var itemsInSession = this.Session["Cart"] as List<string> ?? new List<string>();

            // add all products currently not in session
            var filteredProducts = products.Where(item => !itemsInSession.Contains(item));

            // Add additional filters here
            // filter product list for home users
            if (User.IsInRole("Home"))
            {
                filteredProducts = filteredProducts.Where(item => item.Contains("Home"));
            }

            return View(filteredProducts);
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Add(string selectedItem)
        {
            if (selectedItem != null)
            {
                List<string> cart = this.Session["Cart"] as List<string> ?? new List<string>();
                cart.Add(selectedItem);
                Session["Cart"] = cart;
            }

            return RedirectToAction("Index");
        }

        public ActionResult Checkout()
        {
            var itemsInSession = this.Session["Cart"] as List<string> ?? new List<string>();
            return View(itemsInSession);
        }

        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Remove(string selectedItem)
        {
            if (selectedItem != null)
            {
                var itemsInSession = this.Session["Cart"] as List<string>;
                if (itemsInSession != null)
                {
                    itemsInSession.Remove(selectedItem);
                }
            }

            return RedirectToAction("Checkout");
        }
    }
}