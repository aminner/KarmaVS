using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using devcoach.Tools.Properties;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System.Drawing.Imaging;

namespace devcoach.Tools
{
  /// <summary>
  /// This is the class that implements the package exposed by this assembly.
  /// The minimum requirement for a class to be considered a valid package for Visual Studio
  /// is to implement the IVsPackage interface and register itself with the shell.
  /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
  /// to do it: it derives from the Package class that provides the implementation of the 
  /// IVsPackage interface and uses the registration attributes defined in the framework to 
  /// register itself and its components with the shell.
  /// </summary>
  // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
  // a package.
  [ProvideAutoLoad(UIContextGuids.SolutionExists)]
  [PackageRegistration(UseManagedResourcesOnly = true)]
  // This attribute is used to register the information needed to show this package
  // in the Help/About dialog of Visual Studio.
  [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
  // This attribute is needed to let the shell know that this package exposes some menus.
  [ProvideMenuResource("Menus.ctmenu", 1)]
  [Guid(GuidList.guidKarmaVsPkgString)]
  public sealed class KarmaVsPackage : Package
  {
    private static readonly object _s_applicationLock = new object();
    public static DTE2 Application { get; private set; }
    private static OutputWindow _outputWindow;
    private static OutputWindowPane _karmaOutputWindowPane;
    private static System.Diagnostics.Process _karmaProcess;
    private static System.Diagnostics.Process _webServerProcess;
    private static Events _events;
    private static DTEEvents _dteEvents;
    private static DocumentEvents _documentEvents;
      private static DateTimeOffset _processTime;
      private int errors;
      private bool Enabled;
      private DisplayType _displayType;
      private enum DisplayType
      {
          Success,
          Failure,
          Running,
          BuildError
      };

      private IVsStatusbar _statusBar;
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

    #region Initialize()
    protected override void Initialize()
    {
      lock (_s_applicationLock)
      {
        Application = (DTE2)GetService(typeof(SDTE));
        _events = Application.Application.Events;
        _dteEvents = _events.DTEEvents;
        _documentEvents = _events.DocumentEvents;

        var win =
          Application.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
        _outputWindow = win.Object as OutputWindow;
        if (_outputWindow != null)
        {
          _karmaOutputWindowPane =
            _outputWindow.OutputWindowPanes.Add("Karma");
        }
        _dteEvents.OnBeginShutdown += ShutdownKarma;
        _events.SolutionEvents.Opened += SolutionEventsOpened;

      }

      base.Initialize();
        Enabled = false;
      // Add our command handlers for menu (commands must exist in .vsct file)
      var mcs =
        GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

      //var mcs2 =
      //GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
      if (mcs == null) return;

      // Create the command for the menu item.
      var toggleRun =
        new CommandID(
          GuidList.guidKarmaVsUnitCmdSet,
          (int)PkgCmdIDList.cmdidToggleKarmaVsUnit
        );
        
         var karmaOptions =
        new CommandID(
          GuidList.guidKarmaVsUnitCmdSet,
          (int)PkgCmdIDList.cmdidOptionsKarmaVsUnit
        );
      var menuItem1 = new MenuCommand(KarmaVsUnit, toggleRun);
        var menuItem2 = new MenuCommand(LaunchOptions, karmaOptions);
      mcs.AddCommand(menuItem1);
      mcs.AddCommand(menuItem2);
      _statusBar= (IVsStatusbar)GetService(typeof(SVsStatusbar));
    }
    #endregion

    #region SolutionEventsOpened()
    private void SolutionEventsOpened()
    {
      RunKarmaVs();
    }
    #endregion

    #region RunKarmaVs()
    private void RunKarmaVs(string config = "unit")
    {
        Enabled = true;

        BackgroundWorker worker = new BackgroundWorker();
        worker.DoWork += delegate { SetDisplay(); };
        worker.RunWorkerAsync();
       _displayType = DisplayType.Running;
       errors = 0;
       _karmaOutputWindowPane.Clear();

      ShutdownKarma();

      var nodeFilePath = GetNodeJsPath();
      if (nodeFilePath == null)
      {
        _karmaOutputWindowPane.OutputString(
            "ERROR: Node not found. Download and " +
            "install from: http://www.nodejs.org");
        _karmaOutputWindowPane.OutputString(Environment.NewLine);
        _displayType = DisplayType.BuildError;
        return;
      }

      _karmaOutputWindowPane.OutputString(
          "INFO: Node installation found: " + nodeFilePath);
      _karmaOutputWindowPane.OutputString(Environment.NewLine);

      var karmaFilePath = GetKarmaPath();
      if (karmaFilePath == null)
      {
        _karmaOutputWindowPane.OutputString(
            "ERROR: Karma was not found. Run \"npm install -g karma\"...");
        _karmaOutputWindowPane.OutputString(Environment.NewLine);
        _displayType = DisplayType.BuildError;
        return;
      }

      _karmaOutputWindowPane.OutputString(
          "INFO: Karma installation found: " + karmaFilePath);
      _karmaOutputWindowPane.OutputString(Environment.NewLine);

      var chromePath = GetChromePath();
      if (chromePath != null)
      {
        _karmaOutputWindowPane.OutputString(
            "INFO: Found Google Chrome : " + chromePath);
        _karmaOutputWindowPane.OutputString(Environment.NewLine);
        Environment.SetEnvironmentVariable("CHROME_BIN", chromePath);
      }

      var mozillaPath = GetMozillaPath();
      if (mozillaPath != null)
      {
        _karmaOutputWindowPane.OutputString(
            "INFO: Found Mozilla Firefox: " + mozillaPath);
        _karmaOutputWindowPane.OutputString(Environment.NewLine);
        Environment.SetEnvironmentVariable("FIREFOX_BIN", mozillaPath);
      }

      string karmaConfigFilePath = null;
      string projectDir = null;
        if (Settings.Default.KarmaConfigType == (int) KarmaVsStaticClass.KarmaConfigType.Default)
        {
            karmaConfigFilePath = Path.Combine(projectDir, "karma." + config + ".conf.js");
        }
        else
        {
            karmaConfigFilePath = Settings.Default.KarmaConfigLocation;
        }

        const string webApplication = "{349C5851-65DF-11DA-9384-00065B846F21}";
        const string webSite = "{E24C65DC-7377-472B-9ABA-BC803B73C61A}";
        const string testProject = "{3AC096D0-A1C2-E12C-1390-A8335801FDAB}";
        const string appBuilder = "{070BCB52-5A75-4F8C-A973-144AF0EAFCC9}";
        if (Application.Solution.Projects.Count < 1)
        {
            _karmaOutputWindowPane.OutputString("ERROR: No projects loaded");
            _karmaOutputWindowPane.OutputString(Environment.NewLine);
            _displayType = DisplayType.BuildError;
            return;
        }

        var projects = GetProjects();
        Project karmaProject = null;
        foreach (var project in projects)
        {
            try
            {
                var projectGuids = project.GetProjectTypeGuids(GetService);

                _karmaOutputWindowPane.OutputString(
                    "DEBUG: project '" + project.Name + "' found; GUIDs: " + projectGuids);

                if (projectGuids.Contains(webApplication) ||
                    projectGuids.Contains(webSite) ||
                    projectGuids.Contains(testProject) ||
                    projectGuids.Contains(appBuilder))
                {
                    _karmaOutputWindowPane.OutputString(
                        "INFO: Web / Test project found: " + project.Name);
                    _karmaOutputWindowPane.OutputString(Environment.NewLine);

                    projectDir =
                        Path.GetDirectoryName(project.FileName);

                    if (string.IsNullOrWhiteSpace(projectDir))
                    {
                        try
                        {
                            var fullPath = project.Properties.Item("FullPath");
                            projectDir = fullPath.Value.ToString();
                        }
                        catch (ArgumentException)
                        {

                        }
                    }

                    if (File.Exists(karmaConfigFilePath))
                    {
                        _karmaOutputWindowPane.OutputString(
                            "INFO: Configuration found: " + karmaConfigFilePath);
                        _karmaOutputWindowPane.OutputString(Environment.NewLine);
                        karmaProject = project;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }
        if (karmaProject == null ||
            projectDir == null)
        {
            _karmaOutputWindowPane.OutputString(
                "INFO: No web project found with a file " +
                "named \"karma." + config + ".conf.js\" " +
                "in the root directory. Please check that one exists, or set the location of the config file in the options menu.");
            _karmaOutputWindowPane.OutputString(Environment.NewLine);
            _displayType = DisplayType.BuildError;
            return;
        }
        var nodeServerFilePath = Path.Combine(projectDir, "server.js");

      if (File.Exists(nodeServerFilePath))
      {
        _karmaOutputWindowPane.OutputString("INFO: server.js found.");
        _karmaOutputWindowPane.OutputString(Environment.NewLine);

        _webServerProcess =
            new System.Diagnostics.Process
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
        _karmaOutputWindowPane.OutputString(
            "INFO: Starting node server...");
        _karmaOutputWindowPane.OutputString(Environment.NewLine);
        _webServerProcess.Start();
      }


      _karmaProcess =
          new System.Diagnostics.Process
          {
            StartInfo =
            {
              CreateNoWindow = true,
              FileName = nodeFilePath,
              Arguments =
                  "\"" +
                  karmaFilePath +
                  "\" start \"" +
                  karmaConfigFilePath +
                  "\"",
              RedirectStandardOutput = true,
              RedirectStandardError = true,
              UseShellExecute = false,
              WindowStyle = ProcessWindowStyle.Hidden,
            },
          };
      _karmaProcess.ErrorDataReceived += OutputReceived;
      _karmaProcess.OutputDataReceived += OutputReceived;
      _karmaProcess.Exited += KarmaProcessOnExited;
      _karmaProcess.EnableRaisingEvents = true;


      try
      {
        _karmaOutputWindowPane.OutputString("INFO: Starting karma server...");
        _karmaOutputWindowPane.OutputString(Environment.NewLine);
        _karmaProcess.Start();
          _karmaProcess.BeginOutputReadLine();
       var tmp =   _karmaOutputWindowPane.Collection.DTE;

      }
      catch (Exception ex)
      {
        _karmaOutputWindowPane.OutputString("ERROR: " + ex);
        _karmaOutputWindowPane.OutputString(Environment.NewLine);
        _displayType = DisplayType.BuildError;
      }
    }

    #endregion

    #region KarmaProcessOnExited()
    private void KarmaProcessOnExited(object sender, EventArgs eventArgs)
    {
      ShutdownKarma();
    }
    #endregion

    #region GetProjects()
    public static IList<Project> GetProjects()
    {
      var projects = Application.Solution.Projects;
      var list = new List<Project>();
      var item = projects.GetEnumerator();
      while (item.MoveNext())
      {
        var project = item.Current as Project;
        if (project == null)
        {
          continue;
        }

        if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
        {
          list.AddRange(GetSolutionFolderProjects(project));
        }
        else
        {
          list.Add(project);
        }
      }

      return list;
    } 
    #endregion

    #region GetSolutionFolderProjects()
    private static IEnumerable<Project> GetSolutionFolderProjects(
      Project solutionFolder)
    {
      var list = new List<Project>();
      for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
      {
        var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
        if (subProject == null)
        {
          continue;
        }
        if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
        {
          list.AddRange(GetSolutionFolderProjects(subProject));
        }
        else
        {
          list.Add(subProject);
        }
      }

      return list;
    } 
    #endregion

    #region ShutdownKarma()
    private void ShutdownKarma()
    {
      try
      {
        foreach (System.Diagnostics.Process proc in
            System.Diagnostics.Process.GetProcessesByName("phantomjs"))
        {
          _karmaOutputWindowPane.OutputString("KILL: phantomjs");
          _karmaOutputWindowPane.OutputString(Environment.NewLine);
          proc.Kill();
        }
      }
      catch
      {
      }

      try
      {
        foreach (System.Diagnostics.Process proc in
            System.Diagnostics.Process.GetProcessesByName("phantomjs.exe"))
        {
          _karmaOutputWindowPane.OutputString("KILL: phantomjs.exe");
          _karmaOutputWindowPane.OutputString(Environment.NewLine);
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
    #endregion

    #region GetNodeJsPath()
    private string GetNodeJsPath()
    {
      using (var softwareKey = Registry.CurrentUser.OpenSubKey("Software"))
      {
        if (softwareKey == null) return null;
        using (var nodeJsKey = softwareKey.OpenSubKey("Node.js"))
        {
          if (nodeJsKey == null) return null;
          var nodeJsFilePath = Path.Combine(
            (string)nodeJsKey.GetValue("InstallPath"),
            "node.exe");
          if (!File.Exists(nodeJsFilePath))
          {
            return null;
          }
          return nodeJsFilePath;
        }
      }
    }
    #endregion

    #region GetKarmaPath()
    private string GetKarmaPath()
    {
      var karmaFilePath =
        Path.Combine(
          Environment.GetFolderPath(
            Environment.SpecialFolder.ApplicationData),
          "npm\\node_modules\\karma\\bin\\karma");

      if (File.Exists(karmaFilePath))
      {
        return karmaFilePath;
      }
      return null;
    }
    #endregion

    #region GetChromePath()
    private string GetChromePath()
    {
      var chromeFilePath =
        Path.Combine(
          Environment.GetFolderPath(
            Environment.SpecialFolder.ProgramFiles),
          "Google\\Chrome\\Application\\chrome.exe");
      if (File.Exists(chromeFilePath))
      {
        return chromeFilePath;
      }

      chromeFilePath =
        Path.Combine(
          Environment.GetFolderPath(
            Environment.SpecialFolder.ProgramFilesX86),
          "Google\\Chrome\\Application\\chrome.exe");

      if (File.Exists(chromeFilePath))
      {
        return chromeFilePath;
      }

      chromeFilePath =
        Path.Combine(
          Environment.GetFolderPath(
            Environment.SpecialFolder.LocalApplicationData),
          "Google\\Chrome\\Application\\chrome.exe");

      if (File.Exists(chromeFilePath))
      {
        return chromeFilePath;
      }
      return null;
    }
    #endregion

    #region GetMozillaPath()
    private string GetMozillaPath()
    {
      var mozillaFilePath =
          Path.Combine(
              Environment.GetFolderPath(
                  Environment.SpecialFolder.ProgramFiles),
              "Mozilla Firefox\\firefox.exe");
      if (File.Exists(mozillaFilePath))
      {
        return mozillaFilePath;
      }
      mozillaFilePath =
          Path.Combine(
              Environment.GetFolderPath(
                  Environment.SpecialFolder.ProgramFilesX86),
              "Mozilla Firefox\\firefox.exe");
      if (File.Exists(mozillaFilePath))
      {
        return mozillaFilePath;
      }
      mozillaFilePath =
       Path.Combine(
         Environment.GetFolderPath(
           Environment.SpecialFolder.LocalApplicationData),
         "Mozilla Firefox\\firefox.exe");

      if (File.Exists(mozillaFilePath))
      {
        return mozillaFilePath;
      }
      return null;
    }
    #endregion

    #region KarmaVsUnit()

      private void LaunchOptions(object sender, EventArgs e)
      {
          KarmaOptionsForm optionsForm = new KarmaOptionsForm();
          optionsForm.ShowDialog();
      }
    /// <summary>
    /// This function is the callback used to execute a command when the a menu item is clicked.
    /// See the Initialize method to see how the menu item is associated to this function using
    /// the OleMenuCommandService service and the MenuCommand class.
    /// </summary>
    private void KarmaVsUnit(object sender, EventArgs e)
    {
       
      if (_karmaProcess != null)
      {
        ShutdownKarma();

        _karmaOutputWindowPane.Clear();
        _karmaOutputWindowPane.OutputString("INFO: Karma has shut down!");
        _karmaOutputWindowPane.OutputString(Environment.NewLine);
        Enabled = false;
        return;
      }
    
      RunKarmaVs("unit");
    }
    #endregion

    #region Display()

      private void SetDisplay()
      {
          var currentStatus = _displayType;

          var tmp = new Bitmap(25, 25);
          object icon = tmp.GetHbitmap();
          while (Enabled)
          {
              try
              {
                  if (_karmaProcess!=null && DateTimeOffset.Now - _processTime > TimeSpan.FromSeconds(1))
                  {
                       _displayType = (errors > 0) ? DisplayType.Failure : DisplayType.Success;
                  }
              }
              catch (Exception){}

              if (currentStatus != _displayType)
              {
                  if (_displayType == DisplayType.Running)
                  {
                       errors = 0;
                  }
                  currentStatus = _displayType;
                  int frozen;
                  _statusBar.IsFrozen(out frozen);
                  if (frozen == 0)
                  {
                       _statusBar.Animation(0, ref icon);
                       var image =GetDisplayObject(currentStatus);
                      //var tmp = ConvertToRGB(image);
                       
                       icon = ConvertToRGB(image).GetHbitmap();
                      _statusBar.Animation(1, ref icon);
                  }
              }
          }
          _statusBar.Animation(0, ref icon);
      }

      public static Bitmap GetInvertedBitamp(IVsUIShell5 shell5, Bitmap inputBitmap, Color transparentColor,
          UInt32 backgroundColor)
      {
          var outputBitmap = new Bitmap(inputBitmap);
          try
          {
              outputBitmap.MakeTransparent(transparentColor);

              var rect = new Rectangle(0, 0, outputBitmap.Width, outputBitmap.Height);

              var bitmapData = outputBitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadWrite, outputBitmap.PixelFormat);
              var sourcePointer = bitmapData.Scan0;
              var length = (Math.Abs(bitmapData.Stride)*outputBitmap.Height);
              var outputBytes = new Byte[length - 1];
              Marshal.Copy(sourcePointer, outputBytes, 0, length);

              shell5.ThemeDIBits((UInt32) outputBytes.Length, outputBytes, (UInt32) outputBitmap.Width,
                  (UInt32) outputBitmap.Height, true, backgroundColor);
              Marshal.Copy(outputBytes, 0, sourcePointer, length);
              outputBitmap.UnlockBits(bitmapData);
          }
          catch (Exception ex)
          {

          }
          return outputBitmap;
      }

      public static Bitmap ConvertToRGB(Bitmap original)
      {
          var shell5 = GetIVsUIShell5();

          var backgroundColor = shell5.GetThemedColor(new System.Guid("DC0AF70E-5097-4DD3-9983-5A98C3A19942"), "ToolWindowBackground", 0);
          //var outputBitmap = GetInvertedBitmap(shell5, inputBitmap, almostGreenColor, backgroundColor);
          Bitmap newImage = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb);
          newImage.SetResolution(original.HorizontalResolution,
                                 original.VerticalResolution);
          Graphics g = Graphics.FromImage(newImage);
          g.Clear(ColorTranslator.FromWin32(Convert.ToInt32(backgroundColor & 0xffffff)));
          g.DrawImageUnscaled(original, 0, 0);
          g.Dispose();
          return newImage;
      }
      private static Microsoft.VisualStudio.Shell.Interop.IVsUIShell5 GetIVsUIShell5()
      {

          Microsoft.VisualStudio.OLE.Interop.IServiceProvider sp = null;
          Type SVsUIShellType = null;
          int hr = 0;
          IntPtr serviceIntPtr;
          Microsoft.VisualStudio.Shell.Interop.IVsUIShell5 shell5 = null;
          object serviceObject = null;

          sp = (Microsoft.VisualStudio.OLE.Interop.IServiceProvider)Application;

          SVsUIShellType = typeof(Microsoft.VisualStudio.Shell.Interop.SVsUIShell);

          hr = sp.QueryService(SVsUIShellType.GUID, SVsUIShellType.GUID, out serviceIntPtr);

          if (hr != 0)
          {
              System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hr);
          }
          else if (!serviceIntPtr.Equals(IntPtr.Zero))
          {
              serviceObject = System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(serviceIntPtr);

              shell5 = (Microsoft.VisualStudio.Shell.Interop.IVsUIShell5)serviceObject;

              System.Runtime.InteropServices.Marshal.Release(serviceIntPtr);
          }
          return shell5;
      }
      private Bitmap GetDisplayObject(DisplayType type)
      {
          switch (type)
          {
              case DisplayType.BuildError:
                  return Resources._1420674003_001_11;
              case DisplayType.Failure:
                  return Resources._1420678809_001_19;
              case DisplayType.Running:
                  return Resources._1420678868_001_25;
              case DisplayType.Success:
                  return Resources._1420678794_001_18;
              default:
                  return null;
          }
      }
    #endregion
    #region OutputReceived()
    private void OutputReceived(
          object sender,
          DataReceivedEventArgs dataReceivedEventArgs)
    {
      try
      {
          _processTime = DateTimeOffset.Now;
          _displayType = DisplayType.Running;
          if (dataReceivedEventArgs.Data.ToLower().Contains("failed"))
          {
              errors++;
          }
        _karmaOutputWindowPane.Activate();
        _karmaOutputWindowPane.OutputString(
          FixData(dataReceivedEventArgs.Data));
        _karmaOutputWindowPane.OutputString(
          Environment.NewLine);
      }
      catch { }
    }
    #endregion

    #region FixData()
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
    #endregion
  }
}
