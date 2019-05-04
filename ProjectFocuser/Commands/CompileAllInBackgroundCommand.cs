using EnvDTE;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace ProjectFocuser.Commands
{
    internal sealed class CompileAllInBackgroundCommand : CommandBase
    {
        public override int CommandId => 0x0500;

        private CompileAllInBackgroundCommand(Package package) : base(package)
        {
            AddCommandToMenu(this.MenuItemCallback);
        }

        public static CompileAllInBackgroundCommand Instance
        {
            get;
            private set;
        }

        public static void Initialize(Package package)
        {
            Instance = new CompileAllInBackgroundCommand(package);
        }

        public void MenuItemCallback(object sender, EventArgs e)
        {
            if (ShowErrorMessageAndReturnTrueIfNoSolutionOpen())
            {
                return;
            }

            var componentModel = Package.GetGlobalService(typeof(SComponentModel)) as IComponentModel;
            var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
            var slnPath = dte.Solution.FileName;

            IVsOutputWindowPane customPane = DteUtil.GetThisExtensionOutputPane();

            customPane.OutputStringThreadSafe($"Starting full compilation of {slnPath}\r\n");

            IRoslynSolutionAnalysis roslyn = new RoslynSolutionAnalysis();
            roslyn.CompileFullSolutionInBackgroundAndReportErrors(slnPath, (message) => customPane.OutputStringThreadSafe(message));
        }


    }
}