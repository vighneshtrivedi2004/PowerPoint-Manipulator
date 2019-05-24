using ManipulatorHelper;
using System;

namespace PowerPoint_Manipulator
{
    class Program
    {
        static void Main(string[] args)
        {
            string templatePresentation_FilePath = @System.Configuration.ConfigurationManager.AppSettings["templatePresentation_FilePath"];
            string presentationWithData_FilePath = @System.Configuration.ConfigurationManager.AppSettings["presentationWithData_FilePath"];
            try
            {
                //Copy template ppt file to new ppt
                Console.WriteLine("\nCopying template file...");
                Manipulator.CopyPowerPointPresentation(templatePresentation_FilePath, presentationWithData_FilePath);

                Console.WriteLine("\nManipulating PowerPoint File with content...");
                TestData testData = Manipulator.GetTestData();
                Manipulator.ManipulatePowerPointPresentationWithContent(presentationWithData_FilePath, testData);
                Console.WriteLine("\nThe new Powerpoint file with content has been saved successfully at\n" + presentationWithData_FilePath);
                Console.WriteLine("\nPress any key to continue...");

                Console.ReadKey();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
