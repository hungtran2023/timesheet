using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIS.Data.Model
{
    public class MessageContentModel
    {
        public int id { get; set; }

        public string Title { get; set; }

        public string Description { get; set; }

        public DateTime CreateDate { get;set; }

        public DateTime UpdateDate { get;set; }

        public bool IsActive { get; set; }

    }
}
