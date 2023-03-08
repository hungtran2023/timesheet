using System;

namespace AIS.Domain.Email
{
    public class EmailModel
    {
        public int ManagerId { get; set; }
        public int RequesterId { get; set; }
        public String Note { get; set; }
        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
        public String Template { get; set; }
        public int? CC { get; set; }
        public int? BCC { get; set; }
        public String LinkAprroveRequestForManager { get; set; }
        public bool IsReverse { get; set; }

        public String ServerPath { get; set; }

        public String APK7Character { get; set; }

        public string TimeClose { get; set; }
        public EmailModel()
        {
            IsReverse = false;
        }
    }
}
