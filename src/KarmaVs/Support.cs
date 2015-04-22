using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.Win32;
using System.IO;
using devcoach.Tools.Properties;

namespace devcoach.Tools
{
    public static class Support
    {
        public static string GetKarmaConfigPath(string config = "unit")
        {
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
            return karmaConfigFilePath;
        }
        public static string GetNodeJsPath()
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
         public static string GetChromePath()
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

    public static string GetMozillaPath()
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
        public static string GetKarmaPath()
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
        public static IList<Project> GetProjects()
        {
            var projects = KarmaVsPackage.Application.Solution.Projects;
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
    }
}
