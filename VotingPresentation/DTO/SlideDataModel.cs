using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VotingPresentation.DTO
{
   public class SlideDataModel
    {
        public List<VotingResponseModel> VotingResponseList { get; set; }
        public List<AnswerModel> SlideOptionList { get; set; }
    }
}
