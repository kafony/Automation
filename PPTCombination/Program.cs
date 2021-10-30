using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.PowerPointApi.Tools;
using System;

namespace PPTCombination
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello world");

            var app = new Application();
            var utils = new CommonUtils(app);

            var presentation = app.Presentations.Add(MsoTriState.msoTrue);
            presentation.Slides.Add(1, PpSlideLayout.ppLayoutClipArtAndVerticalText);

            var documentFile = utils.File.Combine(Environment.CurrentDirectory, "demo", DocumentFormat.Normal);
            presentation.SaveAs(documentFile);

            app.Quit();
            app.Dispose();
        }
    }
}