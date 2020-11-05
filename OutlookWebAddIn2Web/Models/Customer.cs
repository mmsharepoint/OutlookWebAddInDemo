using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Web;

namespace OutlookWebAddIn2Web.Models
{
  [DataContract]
  public class Customer
  {
    public Customer(string id, string name)
    {
      ID = id;
      Name = name;
    }
    [DataMember]
    public string ID { get; set; }
    [DataMember]
    public string Name { get; set; }

  }
}