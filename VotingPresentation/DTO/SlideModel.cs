using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VotingPresentation.DTO
{
   public class SlideModel
    {
        public string SlideNumber { get; set; }
        public List<Options> Options { get; set; }
    }
    public class Options
    {
        public int OptionNumber { get; set; }
        public decimal OptionPercentage { get; set; }
    }
}
