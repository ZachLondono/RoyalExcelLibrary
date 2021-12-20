using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Application.Common {
    public class DatabaseConfiguration {

        /// <summary>
        /// Connection string to database which stores application configuration
        /// </summary>
        public string AppConfigConnectionString { get; set; }

        /// <summary>
        /// Connection string to database which stores job data
        /// </summary>
        public string JobConnectionString { get; set; }

    }

}
