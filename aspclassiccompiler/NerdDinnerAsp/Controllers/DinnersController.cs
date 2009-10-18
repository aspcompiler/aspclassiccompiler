using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using NerdDinner.Helpers;
using NerdDinner.Models;

namespace NerdDinner.Controllers {

    //
    // ViewModel Classes

    public class DinnerFormViewModel {

        // Properties
        public Dinner     Dinner    { get; private set; }
        public SelectList Countries { get; private set; }

        // Constructor
        public DinnerFormViewModel(Dinner dinner) {
            Dinner = dinner;
            Countries = new SelectList(PhoneValidator.Countries, Dinner.Country);
        }
    }

    //
    // Controller Class

    [HandleError]
    public class DinnersController : Controller {

        IDinnerRepository dinnerRepository;

        //
        // Dependency Injection enabled constructors

        public DinnersController()
            : this(new DinnerRepository()) {
        }

        public DinnersController(IDinnerRepository repository) {
            dinnerRepository = repository;
        }

        //
        // GET: /Dinners/
        //      /Dinners/Page/2

        public ActionResult Index(int? page) {

            const int pageSize = 10;

            var upcomingDinners = dinnerRepository.FindUpcomingDinners();
            var paginatedDinners = new PaginatedList<Dinner>(upcomingDinners, page ?? 0, pageSize);

            return View(paginatedDinners);
        }

        //
        // GET: /Dinners/Details/5

        public ActionResult Details(int id) {

            Dinner dinner = dinnerRepository.GetDinner(id);

            if (dinner == null)
                return View("NotFound");

            return View(dinner);
        }

        //
        // GET: /Dinners/Edit/5

        [Authorize]
        public ActionResult Edit(int id) {

            Dinner dinner = dinnerRepository.GetDinner(id);

            if (!dinner.IsHostedBy(User.Identity.Name))
                return View("InvalidOwner");

            return View(new DinnerFormViewModel(dinner));
        }

        //
        // POST: /Dinners/Edit/5

        [AcceptVerbs(HttpVerbs.Post), Authorize]
        public ActionResult Edit(int id, FormCollection collection) {

            Dinner dinner = dinnerRepository.GetDinner(id);

            if (!dinner.IsHostedBy(User.Identity.Name))
                return View("InvalidOwner");

            try {
                UpdateModel(dinner);

                dinnerRepository.Save();

                return RedirectToAction("Details", new { id=dinner.DinnerID });
            }
            catch {
                ModelState.AddModelErrors(dinner.GetRuleViolations());

                return View(new DinnerFormViewModel(dinner));
            }
        }

        //
        // GET: /Dinners/Create

        [Authorize]
        public ActionResult Create() {

            Dinner dinner = new Dinner() {
                EventDate = DateTime.Now.AddDays(7)
            };

            return View(new DinnerFormViewModel(dinner));
        } 

        //
        // POST: /Dinners/Create

        [AcceptVerbs(HttpVerbs.Post), Authorize]
        public ActionResult Create(Dinner dinner) {

            if (ModelState.IsValid) {

                try {
                    dinner.HostedBy = User.Identity.Name;

                    RSVP rsvp = new RSVP();
                    rsvp.AttendeeName = User.Identity.Name;
                    dinner.RSVPs.Add(rsvp);

                    dinnerRepository.Add(dinner);
                    dinnerRepository.Save();

                    return RedirectToAction("Details", new { id=dinner.DinnerID });
                }
                catch {
                    ModelState.AddModelErrors(dinner.GetRuleViolations());
                }
            }

            return View(new DinnerFormViewModel(dinner));
        }

        //
        // HTTP GET: /Dinners/Delete/1

        [Authorize]
        public ActionResult Delete(int id) {

            Dinner dinner = dinnerRepository.GetDinner(id);

            if (dinner == null)
                return View("NotFound");

            if (!dinner.IsHostedBy(User.Identity.Name))
                return View("InvalidOwner");

            return View(dinner);
        }

        // 
        // HTTP POST: /Dinners/Delete/1

        [AcceptVerbs(HttpVerbs.Post), Authorize]
        public ActionResult Delete(int id, string confirmButton) {

            Dinner dinner = dinnerRepository.GetDinner(id);

            if (dinner == null)
                return View("NotFound");

            if (!dinner.IsHostedBy(User.Identity.Name))
                return View("InvalidOwner");

            dinnerRepository.Delete(dinner);
            dinnerRepository.Save();

            return View("Deleted");
        }
    }
}
