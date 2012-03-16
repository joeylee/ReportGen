using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGen
{
    class NonProjectReport
    {
        public String category
        { get; set; }
        public String project
        { get; set; }
        public String customer
        { get; set; }
        public String report
        { get; set; }
        public String personInCharge
        { get; set; }
        public String team
        { get; set; }
        public String issue
        { get; set; }
        public String reporter
        { get; set; }
        public DateTime date
        { get; set; }

        public const int cCategory          = 2;
        public const int cProject           = 3;
        public const int cCustomer          = 4;
        public const int cReport            = 5;
        public const int cPersonInCharge    = 6;
        public const int cTeam              = 7;
        public const int cIssue             = 8;
    }
}
