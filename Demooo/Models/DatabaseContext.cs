using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace Demooo.Models {
    public class DatabaseContext : DbContext {
        public DatabaseContext() : base("DBConnection") {
        }

        public DbSet<Customers> Customers { get; set; }
    }
}