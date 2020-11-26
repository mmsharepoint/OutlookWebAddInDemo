using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookWebAddIn2Web.Models
{
  public class AttachmentsReq
  {
    public Attachment[] Attachments { get; set; }
    public string MessageID { get; set; }
    public bool IsValid()
    {
      return Attachments != null && Attachments.Length > 0 && !String.IsNullOrEmpty(MessageID);
    }
  }
}