using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmailSending.Models
{
    public class Mail
    {
        public Mail()
        {
            SendingDateTime = DateTime.Now;
        }

        public int MailId { get; set; }

        public string MailTo { get; set; }

        public string MailFrom { get; set; }

        public string MailSubject { get; set; }

        public string MailBody { get; set; }

        public DateTime SendingDateTime { get; set; }
    }
}