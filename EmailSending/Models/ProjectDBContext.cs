using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace EmailSending.Models
{
    public class ProjectDBContext : DbContext 
    {
         public ProjectDBContext() : base("DefaultConnection")
        {

        }

        public DbSet<Mail> mails { get; set; }

        public DbSet<MailModel> MailModels { get; set; }

    }
}