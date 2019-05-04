using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace ProjectFocuser.Commands
{
    internal sealed class LoadAllProjectsCommand : CommandBase
    {
        public override int CommandId => 0x0200;

        private LoadAllProjectsCommand(Package package) : base(package)
        {
            AddCommandToMenu(this.MenuItemCallback);
        }

        public static LoadAllProjectsCommand Instance
        {
            get;
            private set;
        }

        public static void Initialize(Package package)
        {
            Instance = new LoadAllProjectsCommand(package);
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

                var res = solutionService4.ReloadProject(ref guid);
            }
        }
    }
}
