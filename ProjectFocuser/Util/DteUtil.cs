using EnvDTE;
using EnvDTE80;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml;

namespace ProjectFocuser
{
	public static class DteUtil
	{
		public static bool IsCsproj(this Project proj)
		{
			var isCsproj = proj.UniqueName.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase);

			return isCsproj;
		}

		public static bool IsUnloaded(this Project proj)
		{
			var isUnloaded = string.Compare(EnvDTE.Constants.vsProjectKindUnmodeled, proj.Kind, StringComparison.OrdinalIgnoreCase) == 0;

			return isUnloaded;
		}

		public static bool IsUnloaded(this ProjectItem proj)
		{
			var isUnloaded = string.Compare(EnvDTE.Constants.vsProjectKindUnmodeled, proj.Kind, StringComparison.OrdinalIgnoreCase) == 0;

			return isUnloaded;
		}

		public static bool IsFolder(this Project proj)
		{
			var isFolder = proj.Kind == EnvDTE.Constants.vsProjectKindSolutionItems;

			return isFolder;
		}
		public static bool IsFolder(this ProjectItem proj)
		{
			var isFolder = proj.Kind == EnvDTE.Constants.vsProjectKindSolutionItems;

			return isFolder;
		}

		private static void ForEach(this Projects projects, Action<Project> action)
		{
			for (int count = 1; count <= projects.Count; count++)
			{
				var proj = projects.Item(count);

				action(proj);
			}
		}

		public static ProjectItem[] GetProjectItemsRecursively(IVsSolution solutionService, DTE2 dte)
		{
			var allProjects = new List<ProjectItem>();
			dte.Solution.Projects.ForEach((proj) =>
			{
				if (proj.IsFolder())
				{
					GetSolutionFolderProjects(solutionService, proj, allProjects);
				}
				else
				{
					var path = proj.Name; ;
					var item = dte.Solution.FindProjectItem(path);
					//allProjects.Add(item);
				}
			});
			return allProjects.ToArray();
		}

		public static string[] GetSelectedItemNames(DTE dte)
		{
			return dte.SelectedItems
						.Cast<SelectedItem>()
						.Select(item => item.Name)
						.ToArray();
		}

		private static void GetSolutionFolderProjects(IVsSolution solutionService, Project solutionFolder, List<ProjectItem> allProject)
		{
			for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
			{
				var projectItem = solutionFolder.ProjectItems.Item(i);
				if (projectItem == null)
				{
					continue;
				}

				// If this is another solution folder, do a recursive call, otherwise add
				if (projectItem.IsFolder())
				{
					GetSolutionFolderProjects(solutionService, projectItem, allProject);
				}
				else
				{
					allProject.Add(projectItem);
				}
			}
		}

		private static void GetSolutionFolderProjects(IVsSolution solutionService, ProjectItem solutionFolder, List<ProjectItem> allProject)
		{
			for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
			{
				var projectItem = solutionFolder.ProjectItems.Item(i);
				if (projectItem == null)
				{
					continue;
				}

				// If this is another solution folder, do a recursive call, otherwise add
				if (projectItem.IsFolder())
				{
					GetSolutionFolderProjects(solutionService, projectItem, allProject);
				}
				else
				{
					allProject.Add(projectItem);
				}
			}
		}

		public static Guid GetProjectGuid(IVsSolution solutionService, Project proj)
		{
			IVsHierarchy selectedHierarchy;

			ErrorHandler.ThrowOnFailure(solutionService.GetProjectOfUniqueName(proj.UniqueName, out selectedHierarchy));

			var guid = GetProjectGuid(selectedHierarchy);
			return guid;
		}

		public static Guid GetProjectGuid(IVsSolution solutionService, string uniqueName)
		{
			IVsHierarchy selectedHierarchy;

			ErrorHandler.ThrowOnFailure(solutionService.GetProjectOfUniqueName(uniqueName, out selectedHierarchy));

			var guid = GetProjectGuid(selectedHierarchy);
			return guid;
		}

		public static Guid GetProjectGuid(IVsHierarchy projectHierarchy)
		{
			Guid projectGuid;
			int hr;

			hr = projectHierarchy.GetGuidProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_ProjectIDGuid, out projectGuid);
			ErrorHandler.ThrowOnFailure(hr);

			return projectGuid;
		}

		public static void WriteExtensionOutput(string message, int retyrCount = 20)
		{
			try
			{
				IVsOutputWindowPane customPane = GetThisExtensionOutputPane();
				customPane.OutputStringThreadSafe(message);

			}
			catch (InvalidOperationException ex)
			{
				if (retyrCount > 0 && ex.Message.Contains("has not been loaded yet"))
				{
					System.Threading.Tasks.Task.Factory.StartNew(async () =>
					{
						await System.Threading.Tasks.Task.Delay(5000);
						WriteExtensionOutput(message, retyrCount - 1);
					});
				}
			}
		}

		public static IVsOutputWindowPane GetThisExtensionOutputPane()
		{
			var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
			EnvDTE.Window window = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
			IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

			Guid customGuid = new Guid("BC12B8A0-1678-48D5-B8C0-D3B5EB4D9064");
			string customTitle = "ProjectFocuser";

			outWindow.CreatePane(ref customGuid, customTitle, 1, 1);

			IVsOutputWindowPane customPane;

			outWindow.GetPane(ref customGuid, out customPane);
			return customPane;
		}

		internal static void EnsureSelectedProjReferencesAreLoadedCommand(IServiceProvider provider, bool unloadUnusedProjects)
		{
			var dte = Package.GetGlobalService(typeof(DTE)) as DTE;

			foreach (string projectName in DteUtil.GetSelectedItemNames(dte))
			{
				LoadProject(dte, projectName);
			}
		}

		public static void EnsureProjectsLoadedByNames(DTE dte, HashSet<string> allProjectNamesToLoad, bool unloadUnusedProjects)
		{
			IVsSolution4 solutionService4 = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution4;
			var projects = FileParser.GetProjects(dte.Solution.FileName);

			foreach (var project in projects.Values)
			{
				var shouldBeLoaded = allProjectNamesToLoad.Contains(project.Name);

				var guid = project.Guid;

				int res = 0;
				if (shouldBeLoaded)
				{
					res = solutionService4.ReloadProject(ref guid);
				}
				else if (unloadUnusedProjects && !shouldBeLoaded)
				{
					res = solutionService4.UnloadProject(ref guid, (uint)_VSProjectUnloadStatus.UNLOADSTATUS_UnloadedByUser);
				}
			}
		}

		private static Dictionary<string, CsProject> projects = null;
		private static List<Guid> projectGuids = null;

		public static void LoadProject(DTE dte, string projectName)
		{
			IVsSolution4 solution4 = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution4;

			// Set Current Directory
			Directory.SetCurrentDirectory(dte.Solution.FileName.Remove(dte.Solution.FileName.LastIndexOf('\\')));

			// Load Solution and get Project Names & Guids
			projects = FileParser.GetProjects(dte.Solution.FileName);

			// Read All Project Content
			foreach (CsProject project in projects.Values)
			{
				// Read Xml
				project.Data = File.ReadAllText(project.Path);

				// Parse Xml
				project.XmlDocument = new System.Xml.XmlDocument();
				project.XmlDocument.LoadXml(project.Data);
			}

			// Init ProjectGuids
			projectGuids = new List<Guid>();

			// Load Project
			LoadProject(solution4, projectName);

			// Ensure projects are loaded (NOT sure this does anything)
			solution4.EnsureProjectsAreLoaded((uint)projectGuids.Count, projectGuids.ToArray(), 10);

			// Clear ProjectGuids
			projectGuids = null;
		}

		/// <summary>
		/// Load the project identified by projectName.
		/// </summary>
		/// <param name="solution4"></param>
		/// <param name="projectName"></param>
		private static void LoadProject(IVsSolution4 solution4, string projectName)
		{
			var project = projects[projectName];
			var guid = project.Guid;

			// Check ProjectGuid (skip if found)
			if (projectGuids.Contains(guid)) return;

			// Diagnostics
			Debug.WriteLine($"Loading {projectName}");

			// Iterate Project References and load those
			foreach (XmlNode projectReference in project.XmlDocument.GetElementsByTagName("ProjectReference"))
			{
                // Get next Project Name
                projectName = projectReference["Name"]?.InnerText;

                // No Name field? Try Name attribute
                if (projectName == null)
                {
                    projectName = projectReference.Attributes["Name"]?.Value;
                }

                // No Name field? Try Include attribute
                if (projectName == null)
                {
                    projectName = projectReference.Attributes["Include"]?.Value?.Between("\\", ".csproj");
                }

                // Skip Projects that have no Name
                if (projectName == null) continue;

				// Extract Project Name of Reference and call LoadProject recursively
				LoadProject(solution4, projectName);
			}

			// Reload Target Project in Visual Studio
			int res = solution4.ReloadProject(ref guid);

			// Store Guid (to prevent duplicate loads)
			projectGuids.Add(guid);
		}

		public static string Between(this string target, string start, string finish)
		{
			int startIndex = target.LastIndexOf(start) + start.Length;
			int finishIndex = target.LastIndexOf(finish);

			return target.Substring(startIndex, finishIndex - startIndex);
		}
	}
}
