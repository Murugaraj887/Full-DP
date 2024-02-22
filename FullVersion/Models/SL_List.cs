using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FullVersion.Models
{
    public class SL_List
    {
        public int Id
        {
            get;
            set;
        }
        public string SL
        {
            get;
            set;
        }
        public bool Checked
        {
            get;
            set;
        }

        public int Batch_Id
        {
            get;
            set;
        }
        public string Batch
        {
            get;
            set;
        }
        public bool Batch_Checked
        {
            get;
            set;
        }

        public string Type { get; set; }
        public HttpPostedFileBase uploadFile { get; set; }

        public string hidTAB { get; set; }
    }



}