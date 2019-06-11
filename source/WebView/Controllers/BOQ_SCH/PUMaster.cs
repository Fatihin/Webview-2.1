using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WebView.Controllers

{
    class PUMaster
    {
        public string PU_DESC { get; set; }
        public string PU_UOM { get; set; }
        public float PU_INST_PR { get; set; }
        public float PU_MAT_PR { get; set; }
        public SCHMatches SCHReference { get; set; }
    }
}
