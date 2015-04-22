using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EnvDTE;

namespace devcoach.Tools
{
    sealed internal class KarmaCommandLine
    {
        private OutputWindow _outputWindow;
        private static OutputWindowPane _karmaOutputWindowPane;

        readonly Regex _outputDir =
        new Regex(
          @"\[[\d]{1,6}m{1}",
          RegexOptions.Compiled);

        readonly Regex _browserInfo =
            new Regex(
                @"[a-z|A-Z]{4,20}\s([\d]{1,5}\.{0,1})*\s\" +
                @"([a-z|A-Z|\s|\d]{4,20}\)(\:\s|\s|\]\:\s)",
                RegexOptions.Compiled);

        readonly Regex _details =
            new Regex(
              @"\[\d{1,2}[A-Z]{1}|\[\d{1,2}m{1}",
              RegexOptions.Compiled);

        public KarmaCommandLine()
        {
            var win =
            KarmaVsPackage.Application.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            _outputWindow = win.Object as OutputWindow;
            if (_outputWindow != null)
            {
                _karmaOutputWindowPane =
                  _outputWindow.OutputWindowPanes.Add("Karma");
            }
        }

        public void Clear()
        {
            _karmaOutputWindowPane.Clear();
        }

        public void LogComment(string comment)
        {
            _karmaOutputWindowPane.OutputString(comment);
            _karmaOutputWindowPane.OutputString(Environment.NewLine);
        }

        public void OutputReceived(
       object sender,
       DataReceivedEventArgs dataReceivedEventArgs)
        {
            try
            {
                if (dataReceivedEventArgs.Data.ToLower().Contains("failed"))
                {
                    KarmaVsDisplay.KarmaErrors++;
                }
                _karmaOutputWindowPane.Activate();
                _karmaOutputWindowPane.OutputString(
                  FixData(dataReceivedEventArgs.Data));
                _karmaOutputWindowPane.OutputString(
                  Environment.NewLine);
            }
            catch { }
        }
        private string FixData(string data)
        {
            if (data == null) return null;
            data = _outputDir.Replace(data, string.Empty);
            data = _browserInfo.Replace(data, string.Empty);
            data = _details.Replace(data, string.Empty);
            data = data.TrimStart((char)27);
            if (data.StartsWith("xecuted"))
            {
                data = "E" + data;
            }
            return data;
        } 
    }
}
