using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
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
            slides.Add(new Slide { Header = "TestArea1", Questions = new List<string>() { "Question1", "Question2" } });
            slides.Add(new Slide { Header = "TestArea2", Questions = new List<string>() { "Question1", "Question2", "Question3" } });
            slides.Add(new Slide { Header = "TestArea3", Questions = new List<string>() { "Question1", "Question2", "Question3" } });

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

                        for (int i = 0; i < questionsTable.Rows.Count - 1; i++)
                        {
                            questionsTable[0, i + 1].TextFrame.Text = slide.Questions[i];
                            for (int j = 0; j < questionsTable.Columns.Count - 1; j++)
                            {
                                questionsTable[j + 1, 0].TextFrame.Text = testData.Respondents[j];
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
