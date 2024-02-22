using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FullVersion.Controllers
{
    public class Summary
    {
        public string _parameter;
        public double? _H1 = 0;
        public double? _H2 = 0;
        public double? _FY = 0;
        public string _trend;
        public string _blank;
        public string _blank1;
        public double? _M1 = 0;
        public double? _M2 = 0;
        public double? _M3 = 0;
        public double? _M4 = 0;
        public double? _M5 = 0;
        public double? _M6 = 0;
        public double? _M7 = 0;
        public double? _M8 = 0;
        public double? _M9 = 0;
        public double? _M10 = 0;
        public double? _M11 = 0;
        public double? _M12 = 0;
        public double? _PrevH2 = 0;
        public HttpPostedFileBase summaryFile { get; set; }
    }
}