using System.Collections.Generic;
using System.Linq;

namespace PPTCombination
{
    public class Config
    {
        public string SourceFileName { get; set; }
        public string DestFileName { get; set; }

        public string NewFileName { get; set; }


        public IList<Slide> Slides { get; set; } = new List<Slide>();

        public void ConvertPosition()
        {
            if (Slides == null || !Slides.Any())
            {
                return;
            }

            foreach (var slide in Slides)
            {
                if (slide.Shapes == null || !slide.Shapes.Any())
                {
                    continue;
                }

                foreach (var slideShape in slide.Shapes)
                {
                    if (slideShape.SourceLeft.HasValue)
                    {
                        slideShape.SourceLeft = Util.ConvertToPixel(slideShape.SourceLeft.Value);
                    }
                    if (slideShape.SourceTop.HasValue)
                    {
                        slideShape.SourceTop = Util.ConvertToPixel(slideShape.SourceTop.Value);
                    }
                    if (slideShape.DestLeft.HasValue)
                    {
                        slideShape.DestLeft = Util.ConvertToPixel(slideShape.DestLeft.Value);
                    }
                    if (slideShape.DestTop.HasValue)
                    {
                        slideShape.DestTop = Util.ConvertToPixel(slideShape.DestTop.Value);
                    }
                }
            }
        }
    }

    public class Slide
    {
        public int SourceSlideIndex { get; set; }
        public int DestSlideIndex { get; set; }
        public IList<Shape> Shapes { get; set; } = new List<Shape>();

        public Slide()
        { }

        public Slide(int sourceSlideIndex, int destSlideIndex, float? sourceLeft, float? sourceTop, float? destLeft, float? destTop)
        {
            SourceSlideIndex = sourceSlideIndex;
            DestSlideIndex = destSlideIndex;

            if (sourceLeft != null && sourceTop != null)
            {
                Shapes.Add(new Shape(sourceLeft, sourceTop, destLeft, destTop));
            }
        }
    }

    public class Shape
    {
        public float? SourceLeft { get; set; }
        public float? SourceTop { get; set; }

        public float? DestLeft { get; set; }
        public float? DestTop { get; set; }

        public Shape()
        { }

        public Shape(float? sourceLeft, float? sourceTop, float? destLeft, float? destTop)
        {
            SourceLeft = sourceLeft;
            SourceTop = sourceTop;
            DestLeft = destLeft;
            DestTop = destTop;
        }
    }
}
