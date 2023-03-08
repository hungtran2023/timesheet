using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace AIS.Models
{
    public class ImportExcel
    {
        [Required(ErrorMessage = "Please select file")]     
        public HttpPostedFileBase file { get; set; }
    }
}