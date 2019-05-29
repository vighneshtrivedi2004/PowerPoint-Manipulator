using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ManipulatorHelper
{
    public class Manipulator
    {
        /// <summary>
        /// Copy Source PowerPoint presentation to destination presentation.
        /// </summary>
        /// <param name="srcPresentationPath"></param>
        /// <param name="destPresentationPath"></param>
        public static void CopyPowerPointPresentation(string srcPresentationPath, string destPresentationPath)
        {
            Presentation srcPres = new Presentation(@srcPresentationPath);

            srcPres.Save(@destPresentationPath, SaveFormat.Pptx);
        }

        /// <summary>
        /// Get Test Data
        /// </summary>
        /// <returns></returns>
        public static TestData GetTestData()
        {
            TestData testData = new TestData();
            List<string> respondents = new List<string>() { "AB", "AC", "AD", "EF", "EG", "BA", "BD", "BE", "EF", "EG" };
            List<Slide> slides = new List<Slide>();

            List<Question> listQuestions1 = new List<Question>();

            for (int i = 0; i < 22; i++)
            {
                listQuestions1.Add(new Question { Title = "Question " + (i + 1) + ": This is a sample question for the survey", Responses = new int[] { 4, 4, 5, 4, 1, 4, 4, 5, 4, 2 } });
            }

            slides.Add(new Slide { Header = "TestArea1", AverageScore = 4, Questions = listQuestions1 });

            testData.Respondents = respondents;
            testData.Slides = slides;

            return testData;
        }

        /// <summary>
        /// Manipulate PowerPoint Presentation With Content
        /// </summary>
        /// <param name="presentationWithData_FilePath"></param>
        /// <param name="testData"></param>
        public static void ManipulatePowerPointPresentationWithContent(string presentationWithData_FilePath, TestData testData)
        {
            string listRespondents = null;
            int rowSeperatorNo = 3;
            try
            {
                // Creating a presentation instance
                using (Presentation presentationWithData = new Presentation(presentationWithData_FilePath))
                {
                    //Add respondent template slide to new presentation
                    ISlide respondentTemplateSlide = presentationWithData.Slides[0];
                    presentationWithData.Slides.AddClone(respondentTemplateSlide);

                    ISlide lastRespondentSlide = presentationWithData.Slides.Last();

                    int startSlideIndex = lastRespondentSlide.SlideNumber - 1;

                    foreach (var name in testData.Respondents)
                    {
                        listRespondents += name + ", ";
                    }

                    //Remove additional trailing characters
                    listRespondents = listRespondents.Trim();
                    listRespondents = listRespondents.TrimEnd(',');

                    //Modify newly added respondent template slide with content
                    foreach (IShape shp in lastRespondentSlide.Shapes)
                    {
                        if (shp != null)
                        {
                            var text = ((IAutoShape)shp).TextFrame.Text;
                            if (text.StartsWith("List"))
                            {
                                ((IAutoShape)shp).TextFrame.Text = "Respondents with test list:\n" + listRespondents;
                            }
                        }
                    }

                    //Modify area template slide for each Slide 
                    ISlide areaTemplateSlide = presentationWithData.Slides[1];
                    foreach (var slide in testData.Slides)
                    {
                        List<double> listDblColsWidths = new List<double>();
                        List<double> listDblRowsHeights = new List<double>();
                        int rowCount = 0, colCount = 0;

                        presentationWithData.Slides.AddClone(areaTemplateSlide);

                        ISlide lastSlide = presentationWithData.Slides.Last();

                        //Delete all existing tables from area slide
                        for (int k = 0; k < lastSlide.Shapes.Count(); k++)
                        {
                            var shp = lastSlide.Shapes[k];
                            if (shp is ITable) //Delete if shape is table
                            {
                                lastSlide.Shapes.Remove(shp);
                                k = k - 1; //Decrement counter as one shape has been deleted
                            }
                        }

                        // Iterate through shapes to find the placeholder
                        foreach (IShape shp in lastSlide.Shapes)
                        {
                            if (shp.Placeholder != null)
                            {
                                // Change the text of placeholder                           
                                ((IAutoShape)shp).TextFrame.Text = slide.Header;
                            }
                        }

                        //Add Average Table to slide
                        double[] avgTableColWidth = { 50, 23 };
                        double[] avgTableRowHeight = { 4 };
                        ITable avgTable = lastSlide.Shapes.AddTable(50, 100, avgTableColWidth, avgTableRowHeight);
                        avgTable[0, 0].TextFrame.Text = "Average";
                        avgTable[1, 0].TextFrame.Text = Convert.ToString(slide.AverageScore);
                        avgTable[1, 0].FillFormat.FillType = FillType.Solid;
                        avgTable[1, 0].FillFormat.SolidFillColor.Color = Color.Yellow;

                        avgTable.StylePreset = TableStylePreset.NoStyleNoGrid;
                        avgTable.Name = "tblAverage";

                        // setting table cells' font height
                        PortionFormat portionFormat = new PortionFormat();
                        portionFormat.FontHeight = 7;

                        avgTable.SetTextFormat(portionFormat);

                        int emptyRowsCount = (slide.Questions.Count() / rowSeperatorNo);

                        // Add Question table to slide
                        rowCount = slide.Questions.Count() + 1 + emptyRowsCount;
                        colCount = testData.Respondents.Count() + 1;

                        for (int row = 0; row < rowCount; row++)
                        {
                            listDblRowsHeights.Add(4);
                        }

                        listDblColsWidths.Add(200);
                        for (int col = 1; col < colCount; col++)
                        {
                            listDblColsWidths.Add(33);
                        }

                        double[] dblColsWidths = listDblColsWidths.ToArray();
                        double[] dblRowsHeights = listDblRowsHeights.ToArray();
                        ITable questionsTable = lastSlide.Shapes.AddTable(50, 120, dblColsWidths, dblRowsHeights);

                        questionsTable.StylePreset = TableStylePreset.NoStyleNoGrid;
                        questionsTable.Name = "tblQuestions";

                        ITable shpQuestionsTable = (ITable)lastSlide.Shapes.First(x => x.Name == "tblQuestions");

                        // setting table cells' font height                       
                        shpQuestionsTable.SetTextFormat(portionFormat);

                        int rowSeperatorMod = rowSeperatorNo + 1;

                        for (int i = 1, q = 0; i < questionsTable.Rows.Count; i++)
                        {
                            if (i % rowSeperatorMod == 0)
                            {
                                continue;
                            }

                            questionsTable[0, i].TextFrame.Text = slide.Questions[q].Title;
                            for (int j = 0; j < questionsTable.Columns.Count - 1; j++)
                            {
                                questionsTable[j + 1, 0].TextFrame.Text = testData.Respondents[j];
                                questionsTable[j + 1, i].TextFrame.Text = Convert.ToString(slide.Questions[q].Responses[j]);
                            }

                            q++;
                        }
                    }

                    //Delete all template slides
                    for (int i = startSlideIndex - 1; i >= 0; i--)
                    {
                        presentationWithData.Slides.RemoveAt(i);
                    }

                    //Writing the presentation as a PPTX file
                    presentationWithData.Save(presentationWithData_FilePath, SaveFormat.Pptx);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
