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
            List<string> respondents = new List<string>() { "AB", "AC", "AD", "EF", "EG" };
            List<Slide> slides = new List<Slide>();

            List<Question> listQuestions1 = new List<Question>();
            listQuestions1.Add(new Question { Title = "Question1", Responses = new int[] { 4, 4, 5, 4, 1 } });
            listQuestions1.Add(new Question { Title = "Question2", Responses = new int[] { 2, 3, 5, 4, 4 } });
            slides.Add(new Slide { Header = "TestArea1", Questions = listQuestions1 });

            List<Question> listQuestions2 = new List<Question>();
            listQuestions2.Add(new Question { Title = "Question1", Responses = new int[] { 4, 2, 3, 4, 1 } });
            listQuestions2.Add(new Question { Title = "Question2", Responses = new int[] { 1, 5, 3, 4, 1 } });
            listQuestions2.Add(new Question { Title = "Question3", Responses = new int[] { 5, 3, 5, 4, 4 } });
            slides.Add(new Slide { Header = "TestArea2", Questions = listQuestions2 });

            List<Question> listQuestions3 = new List<Question>();
            listQuestions3.Add(new Question { Title = "Question1", Responses = new int[] { 4, 4, 3, 4, 1 } });
            listQuestions3.Add(new Question { Title = "Question2", Responses = new int[] { 4, 4, 4, 4, 1 } });
            listQuestions3.Add(new Question { Title = "Question3", Responses = new int[] { 4, 5, 2, 4, 1 } });
            slides.Add(new Slide { Header = "TestArea3", Questions = listQuestions3 });

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
                        listRespondents += name + "\n";
                    }

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

                        // Add Question table to slide
                        rowCount = slide.Questions.Count() + 1;
                        colCount = testData.Respondents.Count() + 1;

                        for (int row = 0; row < rowCount; row++)
                        {
                            listDblRowsHeights.Add(5);
                        }

                        listDblColsWidths.Add(385);
                        for (int col = 1; col < colCount; col++)
                        {
                            listDblColsWidths.Add(51);
                        }

                        double[] dblColsWidths = listDblColsWidths.ToArray();
                        double[] dblRowsHeights = listDblRowsHeights.ToArray();
                        ITable questionsTable = lastSlide.Shapes.AddTable(56, 100, dblColsWidths, dblRowsHeights);

                        questionsTable.StylePreset = TableStylePreset.NoStyleTableGrid;

                        for (int i = 0; i < questionsTable.Rows.Count - 1; i++)
                        {
                            questionsTable[0, i + 1].TextFrame.Text = slide.Questions[i].Title;
                            for (int j = 0; j < questionsTable.Columns.Count - 1; j++)
                            {
                                questionsTable[j + 1, 0].TextFrame.Text = testData.Respondents[j];
                                questionsTable[j + 1, i + 1].TextFrame.Text = Convert.ToString(slide.Questions[i].Responses[j]);
                            }
                        }

                        // Set border format for each cell
                        foreach (IRow row in questionsTable.Rows)
                        {
                            foreach (ICell cell in row)
                            {
                                cell.BorderTop.FillFormat.FillType = FillType.Solid;
                                cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Gray;

                                cell.BorderBottom.FillFormat.FillType = FillType.Solid;
                                cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Gray;

                                cell.BorderLeft.FillFormat.FillType = FillType.Solid;
                                cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Gray;

                                cell.BorderRight.FillFormat.FillType = FillType.Solid;
                                cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Gray;
                            }
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
