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
        public List<Question> Questions { get; set; }
    }

    public class Question
    {
        public string Title { get; set; }
        public int[] Responses { get; set; }
    }
}
