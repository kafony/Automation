using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi.Enums;
using NetOffice.VBIDEApi.Enums;
using System;
using PPT = NetOffice.PowerPointApi;
using PPTUtil = NetOffice.PowerPointApi.Tools.Contribution;

namespace PPTCombination
{
    public class PPTApplication : IDisposable
    {
        private PPT.Application _application { get; set; }
        private PPT.Presentation _presentation { get; set; }
        private PPTUtil.CommonUtils _utils { get; set; }
        private string _baseaddress { get; set; }
        public PPTApplication(string baseaddress = null)
        {
            if (string.IsNullOrEmpty(baseaddress))
            {
                _baseaddress = Environment.CurrentDirectory;
            }
            else
            {
                _baseaddress = baseaddress;
            }
        }

        public PPT.Presentation GetPresentation(string filename)
        {
            Open(filename);
            return _presentation;
        }

        public void Open(string filename)
        {
            _application = new PPT.Application();
            _utils = new PPTUtil.CommonUtils(_application);

            string documentFile = _utils.File.Combine(_baseaddress, filename, PPTUtil.DocumentFormat.Normal);
            _presentation = _application.Presentations.Open(documentFile);
        }

        public void Save(string filename, PPTUtil.DocumentFormat documentFormat = PPTUtil.DocumentFormat.Normal)
        {
            string documentFile = _utils.File.Combine(_baseaddress, filename, documentFormat);
            _presentation.SaveAs(documentFile);
        }

        public void CreateShape()
        {
            // add a new presentation with one new slide
            var slide = _presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            // add a label
            var label = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 600, 20);
            label.TextFrame.TextRange.Text = "This slide and created Shapes are created by automation tools.";

            // add a line
            slide.Shapes.AddLine(10, 80, 700, 80);

            // add a wordart
            slide.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect9, "This a WordArt", "Arial", 20,
                                           MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 150);

            // add a star
            slide.Shapes.AddShape(MsoAutoShapeType.msoShape24pointStar, 200, 200, 250, 250);
        }

        public void CreateMacro()
        {
            // add a new presentation with one new slide
            var slide = _presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            // add new module and insert macro. the option "Trust access to Visual Basic Project" must be set
            var module = _presentation.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule;
            string macro = string.Format("Sub TestMacro()\r\n   {0}\r\nEnd Sub", "MsgBox \"Click from ppt marco!\"");
            module.InsertLines(1, macro);

            // add button and connect with macro
            var button = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeActionButtonForwardorNext, 100, 100, 200, 200);
            button.ActionSettings[PpMouseActivation.ppMouseClick].AnimateAction = MsoTriState.msoTrue;
            button.ActionSettings[PpMouseActivation.ppMouseClick].Action = PpActionType.ppActionRunMacro;
            button.ActionSettings[PpMouseActivation.ppMouseClick].Run = "TestMacro";
        }

        public void Dispose()
        {
            if (_application != null)
            {
                _application.Quit();
                _application.Dispose();
            }
        }
    }
}
