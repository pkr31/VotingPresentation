using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VotingPresentation.DTO
{
   public class VotingResponseModel
    {
        public int V_ID { get; set; }
        public int UnicoID { get; set; }
        public string Answer { get; set; }
        public string Slide { get; set; }
    }
}
