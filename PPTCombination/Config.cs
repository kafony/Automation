using System.Collections.Generic;

namespace PPTCombination
{
    public class Config
    {
        public string SourceFileName { get; set; }
        public string DestFileName { get; set; }

        public IList<Slide> Slides { get; set; } = new List<Slide>();
    }

    public class Slide
    {
        public int SourceSlideIndex { get; set; }
        public int DestSlideIndex { get; set; }
    }
}
