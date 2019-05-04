using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace ProjectFocuser.Commands
{
	internal sealed class UnloadAllProjectsCommand : CommandBase
	{
		public override int CommandId => 0x0100;

		private UnloadAllProjectsCommand(Package package) : base(package)
		{
			AddCommandToMenu(MenuItemCallback);
		}

		public static UnloadAllProjectsCommand Instance
		{
			get;
			private set;
		}

		public static void Initialize(Package package)
		{
			Instance = new UnloadAllProjectsCommand(package);
		}

		public void MenuItemCallback(object sender, EventArgs e)
		{
			if (ShowErrorMessageAndReturnTrueIfNoSolutionOpen())
			{
				return;
			}

			var dte = Package.GetGlobalService(typeof(DTE)) as DTE;

			IVsSolution4 solutionService4 = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution4;

			var projects = FileParser.GetProjects(dte.Solution.FileName);

			foreach (var proj in projects.Values)
			{
				Guid guid = proj.Guid;
				var res = solutionService4.UnloadProject(ref guid, (uint)_VSProjectUnloadStatus.UNLOADSTATUS_UnloadedByUser);
			}
		}
	}
}
