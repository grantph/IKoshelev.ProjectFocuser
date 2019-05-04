using EnvDTE;
using Microsoft.Build.Locator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Solution = Microsoft.CodeAnalysis.Solution;

namespace ProjectFocuser
{
	public interface IRoslynSolutionAnalysis
	{
		Task<HashSet<string>> GetRecursivelyReferencedProjects(DTE dte, string[] rootProjects);

		HashSet<string> GetProjectsDirectlyReferencing(string solutionFilePath, string[] rootProjectNames);

		System.Threading.Tasks.Task CompileFullSolutionInBackgroundAndReportErrors(string slnPath, Action<string> writeOutput);
	}

	public class RoslynSolutionAnalysis : IRoslynSolutionAnalysis
	{
		private static object MaxVisualStudionVersionRegistrationLock = new object();

		private static bool MaxVisualStudionVersionHasBeenRegistered = false;

		public static void EnsureMaxVisualStudioMsBuildVersionHasBeenRegistered()
		{
			try
			{
				if (MaxVisualStudionVersionHasBeenRegistered)
				{
					return;
				}

				lock (MaxVisualStudionVersionRegistrationLock)
				{
					if (MaxVisualStudionVersionHasBeenRegistered)
					{
						return;
					}

					var maxVisualStudioInstance = MSBuildLocator
														.QueryVisualStudioInstances()
														.OrderByDescending(x => x.Version.Major)
														.First();

					MSBuildLocator.RegisterInstance(maxVisualStudioInstance);

					MaxVisualStudionVersionHasBeenRegistered = true;

					DteUtil.WriteExtensionOutput($"Using msbuild version from Visual Studio version {maxVisualStudioInstance.Version}, " +
										$"path: {maxVisualStudioInstance.MSBuildPath}");
				}
			}
			catch (Exception ex)
			{
				DteUtil.WriteExtensionOutput($"Error during Visual Studio msbuild registration. {ex.Message}");
			}
		}

		public System.Threading.Tasks.Task CompileFullSolutionInBackgroundAndReportErrors(string slnPath, Action<string> writeOutput)
		{
			return System.Threading.Tasks.Task.Factory.StartNew(async () =>
			{
				try
				{
					var workspace = MSBuildWorkspace.Create();

					var solution = await workspace.OpenSolutionAsync(slnPath);

					ProjectDependencyGraph projectGraph = solution.GetProjectDependencyGraph();

					var projectIds = projectGraph.GetTopologicallySortedProjects().ToArray();
					var success = true;
					foreach (ProjectId projectId in projectIds)
					{
						var project = solution.GetProject(projectId);
						writeOutput($"Compiling {project.Name}\r\n");
						Compilation projectCompilation = await project.GetCompilationAsync();
						if (projectCompilation == null)
						{
							writeOutput($"Error, could not get compilation of {project.Name}\r\n");
							continue;
						}

						var diag = projectCompilation.GetDiagnostics().Where(x => x.IsSuppressed == false
																				  && x.Severity == DiagnosticSeverity.Error);

						if (diag.Any())
						{
							success = false;
						}

						foreach (var diagItem in diag)
						{
							writeOutput(diagItem.ToString() + "\r\n");
						}
					}

					if (success)
					{
						writeOutput($"Compilation successful\r\n");
					}
					else
					{
						writeOutput($"Compilation erros found; You can double-click file path in this pane to open it in VS\r\n");
					}
				}
				catch (Exception ex)
				{
					writeOutput($"Error: {ex.Message}\r\n");
				}
			});
		}

		public async Task<HashSet<string>> GetRecursivelyReferencedProjects(DTE dte, string[] rootProjectNames)
		{
			var componentModel = (IComponentModel)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SComponentModel));
			var workspace = componentModel.GetService<Microsoft.VisualStudio.LanguageServices.VisualStudioWorkspace>();

			DteUtil.LoadProject(dte, rootProjectNames[0]);

//var workspace = MSBuildWorkspace.Create();

			//var solution = workspace.OpenSolutionAsync(solutionFilePath).Result;
			var solution = workspace.CurrentSolution;

			var references = new ConcurrentDictionary<string, object>();

			var tasks = rootProjectNames
							.Select(async (projName) =>
							{
								await FillReferencedProjectSetRecursively(solution, projName, references);
							});

			await System.Threading.Tasks.Task.WhenAll(tasks);

			return new HashSet<string>(references.Keys);
		}

		private async System.Threading.Tasks.Task FillReferencedProjectSetRecursively(Solution solution, string projectName, ConcurrentDictionary<string, object> knownReferences)
		{
			var addedFirst = knownReferences.TryAdd(projectName, null);

			if (!addedFirst)
			{
				return;
			}

			var a = solution.Projects.ToArray();

			var project = solution
							.Projects
							.Single(proj => proj.Name == projectName);

			var tasks = project
							.ProjectReferences
							.Select(async (reference) =>
							{
								var refProject = solution
														.Projects
														.Single(proj => proj.Id.Id == reference.ProjectId.Id);

								await FillReferencedProjectSetRecursively(solution, refProject.Name, knownReferences);
							});

			await System.Threading.Tasks.Task.WhenAll(tasks);
		}

		public HashSet<string> GetProjectsDirectlyReferencing(string solutionFilePath, string[] rootProjectNames)
		{
			var workspace = MSBuildWorkspace.Create();

			var solution = workspace.OpenSolutionAsync(solutionFilePath).Result;

			var rootProjectIds = solution.Projects
										.Where(proj => rootProjectNames.Contains(proj.Name))
										.Select(proj => proj.Id.Id)
										.ToArray();

			var referencingProjects = solution
											.Projects
											.Where(proj => proj.ProjectReferences.Any(reference => rootProjectIds.Contains(reference.ProjectId.Id)))
											.Distinct()
											.Select(proj => proj.Name);

			return new HashSet<string>(referencingProjects);
		}

		/*
		public void LoadProject(DTE dte, string projectName)
		{
			IVsSolution4 solutionService4 = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution4;
			var projects = FileParser.GetProjects(dte.Solution.FileName);

			var guidStr = projects[projectName];
			var guid = new Guid(guidStr);

			//int res = solutionService4.ReloadProject(ref guid);
			int res = solutionService4.EnsureProjectIsLoaded(ref guid, 2);
		}
		*/
	}
}
