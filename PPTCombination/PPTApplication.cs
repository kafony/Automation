using System;
using PPT = NetOffice.PowerPointApi;
using PPTUtil = NetOffice.PowerPointApi.Tools;

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

        public void Save(string filename)
        {
            string documentFile = _utils.File.Combine(_baseaddress, filename, PPTUtil.DocumentFormat.Normal);
            _presentation.SaveAs(documentFile);
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
