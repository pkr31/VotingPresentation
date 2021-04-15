using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

using Office = Microsoft.Office.Core;
using VotingPresentation.Shared;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Graph = Microsoft.Office.Interop.Graph;
using Core = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Core;
using VotingPresentation.DTO;

namespace VotingPresentation
{
    public partial class ThisAddIn
    {
        private Helper _helper;
        PowerPoint.Application pptApp;

        private void ThisAddIn_Startup(object sender,EventArgs e)
        {
            _helper = new Helper();
            List <SlideModel> model = _helper.Comparator("258,259", 3);
            //PowerPoint.Presentation ppt = this.Application.Presentations.Add();
            this.Application.Visible = Core.MsoTriState.msoTrue;
            //Create ppt document
            PowerPoint.Presentation ppt = this.Application.Presentations.Add();
            
            //Add a blank slide
            PowerPoint.Slide slide = ppt.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            //Add chart
            //PowerPoint.Shape shape = slide.Shapes.AddChart(Microsoft.Office.Core.XlChartType.xlBarStacked, 100, 100, 500, 300);
           
            Graph.Chart objChart = (Graph.Chart)slide.Shapes.AddOLEObject(150, 150, 480, 320,
                 "MSGraph.Chart.5", "", Core.MsoTriState.msoFalse, "", 0, "",
                 Core.MsoTriState.msoFalse).OLEFormat.Object;

             /*
            //Get the chart
            PowerPoint.Chart chart = slide.Shapes[1].Chart;
            chart.ChartData.Workbook.Application.Visible = false;
            //Create instance from excel workbook to work with chart data
            PowerPoint.ChartData chartData = chart.ChartData;
            Microsoft.Office.Interop.Excel.Workbook dataWorkbook = (Microsoft.Office.Interop.Excel.Workbook)chartData.Workbook;

            //Get the worksheet of chart
            Microsoft.Office.Interop.Excel.Worksheet dataSheet = dataWorkbook.Worksheets[1];

            //Set the range of chart
            Microsoft.Office.Interop.Excel.Range range = dataSheet.Cells.get_Range("A1", "B5");

            //Set the data
            Microsoft.Office.Interop.Excel.ListObject table = dataSheet.ListObjects["Table1"];
            table.Resize(range);
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A2"))).Value = "January";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A3"))).Value = "February";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A4"))).Value = "March";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A5"))).Value = "April";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B2"))).Value = 5;
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B3"))).Value = 15;
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B4"))).Value = 16;
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B5"))).Value = 40;
            */
            //Set the title of the chart
          //  chart.ChartTitle.Text = "Monthly Sales Report";

            //Save the file         
          //  ppt.SaveAs("AddedChart.pptx");

            //foreach (var item in model)
            //{

            ////Add a blank slide
            //PowerPoint.Slide slide = ppt.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            ////Add chart
            //PowerPoint.Shape shape = slide.Shapes.AddChart(Microsoft.Office.Core.XlChartType.xlBarClustered, 100, 100, 500, 300);

            //    //Get the chart
            //    PowerPoint.Chart chart =   slide.Shapes[1].Chart;
            

            ////Create instance from excel workbook to work with chart data
            //PowerPoint.ChartData chartData = chart.ChartData;
            //Microsoft.Office.Interop.Excel.Workbook dataWorkbook = (Microsoft.Office.Interop.Excel.Workbook)chartData.Workbook;

            ////Get the worksheet of chart
            //Microsoft.Office.Interop.Excel.Worksheet dataSheet = dataWorkbook.Worksheets[1];
            //    //dataSheet.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection;
            //    //Set the range of chart

            //    dataSheet = (Microsoft.Office.Interop.Excel.Worksheet)dataWorkbook.Worksheets.get_Item(1);
           
            //}

            //PowerPoint.Application app = new PowerPoint.Application();
            //app.Visible = Core.MsoTriState.msoTrue; // Sure, let's watch the magic as it happens.

            //PowerPoint.Presentation pres = app.Presentations.Add();
            //PowerPoint._Slide objSlide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);

            //PowerPoint.TextRange textRange = objSlide.Shapes[1].TextFrame.TextRange;
            //textRange.Text = "My Chart";
            //textRange.Font.Name = "Comic Sans MS";  // Oh yeah I did
            //textRange.Font.Size = 24;
            //Graph.Chart objChart = (Graph.Chart)objSlide.Shapes.AddOLEObject(150, 150, 480, 320,
            //    "MSGraph.Chart.8", "", Core.MsoTriState.msoFalse, "", 0, "",
            //    Core.MsoTriState.msoFalse).OLEFormat.Object;

            //objChart.ChartType = Graph.XlChartType.xlBarStacked;
            //objChart.Legend.Position = Graph.XlLegendPosition.xlLegendPositionBottom;
            //objChart.HasTitle = true;
            //objChart.ChartTitle.Text = "Sales for Black Programming & Assoc.";  // I'm a regular comedian


         //      PowerPoint.Presentation ppt = this.Application.Presentations.Add();

            //Add a blank slide
    //        PowerPoint.Slide slide = ppt.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            //Add chart
    //        PowerPoint.Shape shape = slide.Shapes.AddChart(Microsoft.Office.Core.XlChartType.xlBarClustered, 100, 100, 500, 300);

            //Get the chart
       //     PowerPoint.Chart chart = slide.Shapes[1].Chart;

            //Create instance from excel workbook to work with chart data
      //      PowerPoint.ChartData chartData = chart.ChartData;
       //     Microsoft.Office.Interop.Excel.Workbook dataWorkbook = (Microsoft.Office.Interop.Excel.Workbook)chartData.Workbook;

            //Get the worksheet of chart
        //    Microsoft.Office.Interop.Excel.Worksheet dataSheet = dataWorkbook.Worksheets[1];
            //dataSheet.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection;
            //Set the range of chart


      //      dataSheet = (Microsoft.Office.Interop.Excel.Worksheet)dataWorkbook.Worksheets.get_Item(1);
            // next 2 lines for split pane in Excel:
            //dataSheet.Application.ActiveWindow.SplitRow = 2;
       //     dataSheet.Application.ActiveWindow.FreezePanes = true;
            //dataSheet.Cells[1, 1] = "Now open the";
            //dataSheet.Cells[2, 1] = "Excel Options window";
            //Microsoft.Office.Interop.Excel.Range range = dataSheet.Cells.get_Range("A2", "B5");

            ////Set the data
            //Microsoft.Office.Interop.Excel.ListObject table = dataSheet.ListObjects["Table1"];
            //table.Resize(range);
            //((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A2"))).Value = "January";
            //((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A3"))).Value = "February";
            //((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A4"))).Value = "March";
            //((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A5"))).Value = "April";
            //((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B2"))).Value = 100;
            //((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B3"))).Value = 250;
            //((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B4"))).Value = 300;
            //((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B5"))).Value = 400;

            //Set the title of the chart
          //  chart.ChartTitle.Text = "Monthly Sales Report";

        }

        private void ThisAddIn_Shutdown(object sender,EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
