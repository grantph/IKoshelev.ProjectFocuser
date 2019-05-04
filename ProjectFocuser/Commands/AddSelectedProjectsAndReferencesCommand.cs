using Microsoft.VisualStudio.Shell;
using System;

namespace ProjectFocuser.Commands
{
    internal sealed class AddSelectedProjectsAndReferencesCommand : CommandBase
    {
        public override int CommandId => 0x0400;

        private AddSelectedProjectsAndReferencesCommand(Package package) : base(package)
        {
            AddCommandToMenu(this.MenuItemCallback);
        }

        public static AddSelectedProjectsAndReferencesCommand Instance
        {
            get;
            private set;
        }

        public static void Initialize(Package package)
        {
            Instance = new AddSelectedProjectsAndReferencesCommand(package);
        }

        public void MenuItemCallback(object sender, EventArgs e)
        {
            if (ShowErrorMessageAndReturnTrueIfNoSolutionOpen())
            {
                return;
            }

            DteUtil.EnsureSelectedProjReferencesAreLoadedCommand(ServiceProvider, false);
        }
    }
}
