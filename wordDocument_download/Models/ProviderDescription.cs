using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace wordDocument_download.Models
{
    public class ProviderDescription
    {
        public string ICD10 { get; set; }
        public string PotentialInaccuracy { get; set; }
        public string DOSLocation { get; set; }
        public string SupportingDoc { get; set; }
    }

}