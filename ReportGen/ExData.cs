using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGen
{
    class ProjectReport
    {
        //private string _project, _record, _plan, _issue;
        public String project
        {
            get;
            set;
        }
        public String record
        {
            get;
            set;
        }
        public String plan
        {
            get;
            set;
        }
        public String issue
        {
            get;
            set;
        }
        public String reporter
        { get; set; }
        public DateTime date
        { get; set; }
    }

}
