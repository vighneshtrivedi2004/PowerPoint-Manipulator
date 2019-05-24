using System.Collections.Generic;

namespace ManipulatorHelper
{
    public class TestData
    {
        public List<string> Respondents { get; set; }
        public List<Slide> Slides { get; set; }
    }

    public class Slide
    {
        public string Header { get; set; }
        public List<string> Questions { get; set; }
    }
}
