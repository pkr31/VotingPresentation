using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VotingPresentation.DTO;

namespace VotingPresentation.Shared
{
    public class Helper
    {
        private DBUtils _dBUtil;

        public List<SlideModel> Comparator(string slideRange ,int finalSlideIndex)
        {
            _dBUtil = new DBUtils();
            List<int> T_ID = slideRange.Split(',').Select(int.Parse).ToList();
            return _dBUtil.GetData(T_ID);
        }
    }
}
