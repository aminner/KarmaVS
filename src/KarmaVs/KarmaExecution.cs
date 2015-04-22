using System;
using System.Diagnostics;
using devcoach.Tools.Properties;
using EnvDTE;
using System.IO;
using Process = System.Diagnostics.Process;

namespace devcoach.Tools
{
    internal sealed class KarmaExecution
    {
        public KarmaVsDisplay _displaySettings { get; private set; }
        public KarmaCommandLine _commandLine { get; private set; }
        public static Process _karmaProcess;
        private static Process _webServerProcess;

        public static DateTimeOffset _processStart;

        public KarmaExecution()
        {
            _displaySettings = new KarmaVsDisplay();
            _commandLine = new KarmaCommandLine();
        }

        public void ShutdownKarma()
        {
            try
            {
                foreach (Process proc in
                    Process.GetProcessesByName("phantomjs"))
                {
                    _commandLine.LogComment("KILL: phantomjs");
                    proc.Kill();
                }
            }
            catch
            {
            }
            if (_karmaProcess != null)
            {
                try
                {
                    _karmaProcess.Kill();
                }
                catch { }
                _karmaProcess = null;
            }
            if (_webServerProcess != null)
            {
                try
                {
                    _webServerProcess.Kill();
                }
                catch { }
                _webServerProcess = null;
            }
        }
        public void StopKarma()
        {
            if (_karmaProcess != null)
            {
                ShutdownKarma();
                _commandLine.Clear();
                _commandLine.LogComment("INFO: Karma has shut down!");
                _displaySettings.Enabled = false;
            }
        }
        internal void StartKarma(string config = "unit")
        {
            _displaySettings.Enabled = true;
            _commandLine.Clear();
            ShutdownKarma();

            var nodeFilePath = Support.GetNodeJsPath();
            if (nodeFilePath == null)
            {
                _commandLine.LogComment(
                    "ERROR: Node not found. Download and " +
                    "install from: http://www.nodejs.org");
                _displaySettings.Type = KarmaVsDisplay.DisplayType.BuildError;
                return;
            }

            _commandLine.LogComment(
                "INFO: Node installation found: " + nodeFilePath);

            var karmaFilePath = Support.GetKarmaPath();
            if (karmaFilePath == null)
            {
                _commandLine.LogComment("ERROR: Karma was not found. Run \"npm install -g karma\"...");
                
                _displaySettings.Type = KarmaVsDisplay.DisplayType.BuildError;
                return;
            }

            _commandLine.LogComment("INFO: Karma installation found: " + karmaFilePath);

            var chromePath = Support.GetChromePath();
            if (chromePath != null)
            {
                _commandLine.LogComment("INFO: Found Google Chrome : " + chromePath);
                Environment.SetEnvironmentVariable("CHROME_BIN", chromePath);
            }

            var mozillaPath = Support.GetMozillaPath();
            if (mozillaPath != null)
            {
                _commandLine.LogComment("INFO: Found Mozilla Firefox: " + mozillaPath);
                Environment.SetEnvironmentVariable("FIREFOX_BIN", mozillaPath);
            }

            if (KarmaVsPackage.Application.Solution.Projects.Count < 1)
            {
                _commandLine.LogComment("ERROR: No projects loaded");
                _displaySettings.Type = KarmaVsDisplay.DisplayType.BuildError;
                return;
            }

            if (KarmaVsPackage.KarmaProject == null ||
                KarmaVsPackage.ProjectDirectory== null)
            {
                _commandLine.LogComment("INFO: No web project found with a file " +
                    "named \"karma." + config + ".conf.js\" " +
                    "in the root directory. Please check that one exists, or set the location of the config file in the options menu.");
                _displaySettings.Type = KarmaVsDisplay.DisplayType.BuildError;
                return;
            }
            var nodeServerFilePath = Path.Combine(KarmaVsPackage.ProjectDirectory, "server.js");

            if (File.Exists(nodeServerFilePath))
            {
                _commandLine.LogComment("INFO: server.js found.");

                _webServerProcess =
                    new Process
                    {
                        StartInfo =
                        {
                            CreateNoWindow = true,
                            FileName = nodeFilePath,
                            Arguments = nodeServerFilePath,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            UseShellExecute = false,
                            WindowStyle = ProcessWindowStyle.Hidden,
                        },
                    };
                _commandLine.LogComment("INFO: Starting node server...");
                _webServerProcess.Start();
            }

            _karmaProcess =
                new Process
                {
                    StartInfo =
                    {
                        CreateNoWindow = true,
                        FileName = nodeFilePath,
                        Arguments =
                            "\"" +
                            karmaFilePath +
                            "\" start \"" +
                            Support.GetKarmaConfigPath() +
                            "\"",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        WindowStyle = ProcessWindowStyle.Hidden,
                    },
                };
            _karmaProcess.ErrorDataReceived += _commandLine.OutputReceived;
            _karmaProcess.OutputDataReceived += _commandLine.OutputReceived;
            _karmaProcess.Exited += KarmaProcessOnExited;
            _karmaProcess.EnableRaisingEvents = true;

            try
            {
                _commandLine.LogComment("INFO: Starting karma server...");
                _karmaProcess.Start();
                _karmaProcess.BeginOutputReadLine();
            }
            catch (Exception ex)
            {
                _commandLine.LogComment("ERROR: " + ex);
                _displaySettings.Type = KarmaVsDisplay.DisplayType.BuildError;
            }
        }
        private void KarmaProcessOnExited(object sender, EventArgs eventArgs)
        {
            ShutdownKarma();
        }
       
    }
}
