using Demooo.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Web;
using System.Web.Mvc;

namespace Demooo.Controllers
{
    public class DemoController : Controller
    {
        // GET: Demo
        public ActionResult Index()
        {
            var model = new ExportViewModel();
            return View(model);
        }

        //Gets the to-do Lists.    
        public JsonResult GetCustomers(string word, int page, int rows, string searchString) {
            //#1 Create Instance of DatabaseContext class for Accessing Database.  
            using (DatabaseContext db = new DatabaseContext()) {
                //#2 Setting Paging  
                int pageIndex = Convert.ToInt32(page) - 1;
                int pageSize = rows;

                //#3 Linq Query to Get Customer   
                var Results = db.Customers.Select(
                    a => new
                    {
                        a.CustomerID,
                        a.CompanyName,
                        a.ContactName,
                        a.ContactTitle,
                        a.City,
                        a.PostalCode,
                        a.Country,
                        a.Phone,
                    });

                //#4 Get Total Row Count  
                int totalRecords = Results.Count();
                var totalPages = (int)Math.Ceiling((float)totalRecords / (float)rows);

                //#5 Setting Sorting  
                //if (sort.ToUpper() == "DESC") {
                //    Results = Results.OrderByDescending(s => s.CustomerID);
                //    Results = Results.Skip(pageIndex * pageSize).Take(pageSize);
                //}
                //else {
                //    Results = Results.OrderBy(s => s.CustomerID);
                //    Results = Results.Skip(pageIndex * pageSize).Take(pageSize);
                //}
                //#6 Setting Search  
                if (!string.IsNullOrEmpty(searchString)) {
                    Results = Results.Where(m => m.Country == searchString);
                }
                //#7 Sending Json Object to View.  
                var jsonData = new
                {
                    total = totalPages,
                    page,
                    records = totalRecords,
                    rows = Results
                };
                return Json(jsonData, JsonRequestBehavior.AllowGet);
            }
        }       
    }
}