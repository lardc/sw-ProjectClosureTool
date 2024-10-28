using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using Newtonsoft.Json;
using System.Threading.Tasks;

namespace ProjectClosureToolMVVM
{
    public partial class ConfProg
    {
        [JsonProperty(nameof(APIKey))]
        public required string APIKey { get; set; }

        [JsonProperty(nameof(myTrelloToken))]
        public required string myTrelloToken { get; set; }

        [JsonProperty(nameof(boardCode))]
        public required string boardCode { get; set; }

        [JsonProperty(nameof(IgnoredLabels))]
        public required List<TrelloObjectLabels> IgnoredLabels { get; set; }
    }
}
