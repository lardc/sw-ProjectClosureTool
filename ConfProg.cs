﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using Newtonsoft.Json;

namespace ProjectClosureToolV2
{
    public partial class ConfProg
    {
        [JsonProperty("APIKey")]
        public string APIKey { get; set; }

        [JsonProperty("myTrelloToken")]
        public string myTrelloToken { get; set; }

        [JsonProperty("boardCode")]
        public string boardCode { get; set; }

        [JsonProperty("IgnoredLabels")]
        public List<TrelloObjectLabels> IgnoredLabels { get; set; }
    }
}