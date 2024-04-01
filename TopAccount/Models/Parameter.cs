using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
namespace TopAccount.Models
{
    public class Parameter
    {
        public Parameter()
        {
            this.SL = new List<SelectListItem>();
            this.From = new List<SelectListItem>();
            this.TO = new List<SelectListItem>();
        }

        public List<SelectListItem> SL { get; set; }
        public List<SelectListItem> From { get; set; }
        public List<SelectListItem> TO { get; set; }

        public string SLId { get; set; }
        public string FromId { get; set; }
        public string ToId { get; set; }
    }
}