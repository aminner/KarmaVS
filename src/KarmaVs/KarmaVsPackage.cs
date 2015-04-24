using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Configuration;
using System.Data.OleDb;
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
using Microsoft.VisualStudio.CommandBars;

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
        public static IVsStatusbar StatusBar;
        public static string ProjectDirectory;
        public static string ProjectGuids;


        private static Events _events;private static DTEEvents _dteEvents;
        public static IMenuCommandService mcs;

        private static KarmaExecution _karmaExecution;

        const string webApplication = "{349C5851-65DF-11DA-9384-00065B846F21}";
        const string webSite = "{E24C65DC-7377-472B-9ABA-BC803B73C61A}";
        const string testProject = "{3AC096D0-A1C2-E12C-1390-A8335801FDAB}";
        const string appBuilder = "{070BCB52-5A75-4F8C-A973-144AF0EAFCC9}";

        #region Initialize()

        private KarmaVsStaticClass.SolutionState SolutionState;
        protected override void Initialize()
        {
            lock (_s_applicationLock)
            {
                Application = (DTE2) GetService(typeof (SDTE));
                _events = Application.Application.Events;
                _dteEvents = _events.DTEEvents;

                _karmaExecution = new KarmaExecution();

                _dteEvents.OnBeginShutdown += _karmaExecution.ShutdownKarma;
                _events.SolutionEvents.Opened += SolutionEventsOpened;
                _events.SolutionEvents.AfterClosing += SolutionEvents_AfterClosing;
            }

            // Add our command handlers for menu (commands must exist in .vsct file)
            mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (mcs == null) return;

            mcs.AddCommand(new MenuCommand(LaunchOptions, new CommandID(
                    GuidList.guidKarmaVsUnitCmdSet,
                    (int)PkgCmdIDList.cmdidOptionsKarmaVsUnit
                    )));

            StatusBar = (IVsStatusbar)GetService(typeof(SVsStatusbar));
            SolutionState = KarmaVsStaticClass.SolutionState.Unloaded;
            base.Initialize();
        }

        void SolutionEvents_AfterClosing()
        {
            SolutionState = KarmaVsStaticClass.SolutionState.Unloaded;
        }

        #endregion

        private void SolutionEventsOpened()
        {
            GetProjects();
            if (Settings.Default.Properties[ProjectGuids] == null)
            {
                var settingsProperty = new SettingsProperty(ProjectGuids);
                Settings.Default.Properties.Add(settingsProperty);
                Settings.Default.Properties[ProjectGuids].Attributes.Add("ConfigLocation", "");
                Settings.Default.Properties[ProjectGuids].Attributes.Add("ConfigType", KarmaVsStaticClass.KarmaConfigType.Default);
            }
            
            SolutionState = KarmaVsStaticClass.SolutionState.Loaded;
            _karmaExecution.StartKarma();
        }

        void toggle_BeforeQueryStatus(object sender, EventArgs e)
        {
            var myCommand = sender as OleMenuCommand;
            if (null != myCommand)
            {
                if (!_karmaExecution._displaySettings.Enabled)
                {
                    myCommand.Text = "Enable";
                }
                else
                {
                    myCommand.Text = "Disable";
                }
            }
        }

        public static void SetMenuStatus()
        {
            if (mcs == null || _karmaExecution == null)
                return;
            var run = mcs.FindCommand(new CommandID(GuidList.guidKarmaVsUnitCmdSet, (int)PkgCmdIDList.cmdidRunTests));
            run.Enabled = _karmaExecution._displaySettings.Enabled;

            var toggle= mcs.FindCommand(new CommandID(GuidList.guidKarmaVsUnitCmdSet, (int)PkgCmdIDList.cmdidToggleKarmaVsUnit)) as OleMenuCommand;;
            if (toggle == null)
               return;
            toggle.Text = _karmaExecution._displaySettings.Enabled ? "Enable" : "Disable";
        }
        private void GetProjects()
        {
            var projects = Support.GetProjects();
            foreach (var project in projects)
            {
                try
                {
                    ProjectGuids = project.GetProjectTypeGuids(GetService);
                    _karmaExecution._commandLine.LogComment("DEBUG: project '" + project.Name + "' found; GUIDs: " + ProjectGuids);
                    if (ProjectGuids.Contains(webApplication) ||
                        ProjectGuids.Contains(webSite) ||
                        ProjectGuids.Contains(testProject) ||
                        ProjectGuids.Contains(appBuilder))
                    {
                        _karmaExecution._commandLine.LogComment("INFO: Web / Test project found: " + project.Name);

                        ProjectDirectory =
                            Path.GetDirectoryName(project.FileName);
                        var karmaConfigFilePath = Support.GetKarmaConfigPath();

                        if (String.IsNullOrWhiteSpace(ProjectDirectory))
                        {
                            try
                            {
                                var fullPath = project.Properties.Item("FullPath");
                                ProjectDirectory = fullPath.Value.ToString();
                            }
                            catch (ArgumentException){}
                        }
                        if (File.Exists(karmaConfigFilePath))
                        {
                            _karmaExecution._commandLine.LogComment("INFO: Configuration found: " + karmaConfigFilePath);
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                }
            }
        }
        private void LaunchOptions(object sender, EventArgs e)
        {
            KarmaOptionsForm optionsForm = new KarmaOptionsForm();
            optionsForm.ShowDialog();
        }
        private void KarmaVsUnitEnable(object sender, EventArgs e)
        {
            var toggle = mcs.FindCommand(new CommandID(GuidList.guidKarmaVsUnitCmdSet, (int)PkgCmdIDList.cmdidToggleKarmaVsUnit)) as OleMenuCommand; ;
            if (toggle == null)
                return;
            if (toggle.Text == "Disable")
            {
                _karmaExecution.StopKarma();
            }
            else
            {
                _karmaExecution.StartKarma();
            }
        }
        private void KarmaVsUnitRun(object sender, EventArgs e)
        {
             _karmaExecution.StartKarma();
        }

    }
}
