using EnvDTE;
using ProjectFocuser.UI;
using Microsoft.VisualStudio.Shell;
using System;

namespace ProjectFocuser.Commands
{
    internal sealed class SaveCurrentSuoCommand : CommandBase
    {
        public override int CommandId => 0x0700;

        private SaveCurrentSuoCommand(Package package) : base(package)
        {
            AddCommandToMenu(this.MenuItemCallback);
        }

        public static SaveCurrentSuoCommand Instance
        {
            get;
            private set;
        }

        public static void Initialize(Package package)
        {
            Instance = new SaveCurrentSuoCommand(package);
        }

        public void MenuItemCallback(object sender, EventArgs e)
        {
            if (ShowErrorMessageAndReturnTrueIfNoSolutionOpen())
            {
                return;
            }

            var dte = Package.GetGlobalService(typeof(DTE)) as DTE;

            var documentationControl = new UI.SaveCurrentSuoDialog();
            documentationControl.DataContext = new SaveCurrentSuoDialogVM(dte.Solution.FileName);
            documentationControl.ShowDialog();
        }
    }
}
