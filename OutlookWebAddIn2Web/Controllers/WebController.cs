using OutlookWebAddIn2Web.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace OutlookWebAddIn2Web.Controllers
{
    public class WebController : ApiController
    {
      private IEnumerable<Customer> data = new List<Customer>() { new Customer("4711", "Customer 1"), 
                                                                  new Customer("4712", "Customer 2"), 
                                                                  new Customer("4713", "Customer 3" ), 
                                                                  new Customer("4714", "Customer 4"), 
                                                                  new Customer("4715", "Customer 5"), 
                                                                  new Customer("4716", "Customer 6"), 
                                                                  new Customer("4717", "Customer 7"), 
                                                                  new Customer("4718", "Customer 8"), 
                                                                  new Customer("4719", "Customer 9"), 
                                                                  new Customer("4720", "Customer 10"), 
                                                                  new Customer("4721", "Customer 11"), 
                                                                  new Customer("4722", "Customer 12") };
    // GET api/<controller>
      public IEnumerable<Customer> Get()
      {
        return data;
      }

    public Customer Get(string id)
    {
      return data.FirstOrDefault(c => c.ID == id);
    }

    // api/<controller>/Search/<querystring>
    [HttpGet]
    public IEnumerable<Customer> Search(string query)
    {
      return data.Where(c => c.Name.Contains(query));
    }
  }
}
