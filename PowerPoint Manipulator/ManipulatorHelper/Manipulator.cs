using Aspose.Slides;
using Aspose.Slides.Export;
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
                    presentationWithData.Slides.AddClone(areaTemplateSlide);

                    ISlide lastSlide = presentationWithData.Slides.Last();

                    // Iterate through shapes to find the placeholder
                    foreach (IShape shp in lastSlide.Shapes)
                    {
                        if (shp.Placeholder != null)
                        {
                            // Change the text of placeholder                           
                            ((IAutoShape)shp).TextFrame.Text = slide.Header;
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
    }
}
