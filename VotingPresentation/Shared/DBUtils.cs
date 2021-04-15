using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VotingPresentation.DTO;

namespace VotingPresentation.Shared
{
    public class DBUtils
    {
        public static OdbcConnection getOdbcConnection()
        {
            string dbfilepath = System.IO.Path.GetFullPath("Data//ComparadaUno[20191015100652886].mdb");
            string connString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" + dbfilepath + ";";
            OdbcConnection conn = new OdbcConnection(connString);
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            return conn;
        }


        public List<SlideModel> GetData(List<int> T_ID)
        {
            try
            {
                //SlideDataModel slideDataModel = new SlideDataModel();
                List<DataTable> dtList = new List<DataTable>();
                //  List<SlideResponse> slideResponse = new List<SlideResponse>();
                List<VotingResponseModel> votingResponse = new List<VotingResponseModel>();
                //List<AnswerModel> slidesOptionList = new List<AnswerModel>();
                var SlideOptions = new Dictionary<string, string>();
                OdbcConnection conn = getOdbcConnection();
                OdbcDataAdapter adapter;
                DataTable table;
                OdbcCommand cmd;
                cmd = new OdbcCommand();
                // cmd.CommandText = "SELECT * FROM ST_Response2 WHERE T_ID in (" + "'" + String.Join("','", T_ID) + "'" + ")";
                cmd.CommandText = "SELECT  T_ID As Slide , TP_Value AS Questions  From ST_TopicPar WHERE TP_Name = 'OptionCount' AND t_id in (" + "'" + String.Join("','", T_ID) + "'" + ")";
                cmd.Connection = conn;
                adapter = new OdbcDataAdapter(cmd);
                table = new DataTable();
                adapter.Fill(table);
                if (table.Rows.Count > 0)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //AnswerModel entity = new AnswerModel();
                        //entity.SlideOptions = new Dictionary<string, string>();
                        //entity.SlideOptions.Add(Convert.ToString(table.Rows[i][0]), Convert.ToString(table.Rows[i][1]));
                        SlideOptions.Add(Convert.ToString(table.Rows[i][0]), Convert.ToString(table.Rows[i][1]));
                    }
                }
                //slideDataModel.SlideOptionList = slidesOptionList;
                cmd = new OdbcCommand();
                cmd.Connection = conn;
                cmd.CommandText = $@"SELECT A.v_id AS V_ID, Max (A.r_id) AS UnicoID,
                                     B.r_result AS Answer, a.t_ID AS Slide
                                     FROM ST_Response AS A LEFT JOIN ST_Response AS B
                                     ON (A.r_id = B.r_id) AND (A.v_id = B.v_id)
                                     WHERE A.t_id In (" + "'" + string.Join("','", T_ID) + "'" + ")" +
                                    @"GROUP BY A.v_id, b.r_result, a.t_id
                                       ORDER BY a.t_id";
                adapter = new OdbcDataAdapter(cmd);
                table = new DataTable();
                adapter.Fill(table);
                if (table.Rows.Count > 0)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        VotingResponseModel entity = new VotingResponseModel();
                        entity.V_ID = Convert.ToInt32(table.Rows[i][0]);
                        entity.UnicoID = Convert.ToInt32(table.Rows[i][1]);
                        entity.Answer = Convert.ToString(table.Rows[i][2]);
                        entity.Slide = Convert.ToString(table.Rows[i][3]);
                        votingResponse.Add(entity);
                    }
                }
                //slideDataModel.VotingResponseList = votingResponse;
                var slideResults = votingResponse.GroupBy(n => new { n.Slide })
            .Select(g => new { Slide = g.Key, Votes = g });

                List<SlideModel> list = new List<SlideModel>();
                foreach (var result in slideResults)
                {
                    var optionResult = result.Votes.GroupBy(n => new { n.Answer })
                    .Select(g => new Options { OptionNumber = Convert.ToInt32(g.Key.Answer.Replace(",", "")), OptionPercentage = g.Count() * 100 / result.Votes.Count() });
                    list.Add(new SlideModel
                    {
                        SlideNumber =
                        result.Slide.Slide,
                        Options = optionResult.Select(x =>
                        new Options
                        {
                            OptionNumber = x.OptionNumber,
                            OptionPercentage = x.OptionPercentage
                        }).ToList()
                    });
                }

                foreach (var slide in list)
                {
                    int length = Convert.ToInt32(SlideOptions[slide.SlideNumber]);
                    for (int i = 1; i <= length; i++)
                    {
                        if (!slide.Options.Any(x => x.OptionNumber == i))
                        {
                            slide.Options.Add(new Options { OptionNumber = i, OptionPercentage = 0 });
                        }
                    }
                    slide.Options = slide.Options.OrderBy(x => x.OptionNumber).ToList();
                }
                return list;
            }
            catch (Exception ex)
            {
                return new List<SlideModel>();
            }
        }

    }
}
