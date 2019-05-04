using Microsoft.VisualStudio.Shell;
using System;

namespace ProjectFocuser.Commands
{
    internal sealed class EnsureOnlySelectedProjReferencesAreLoadedCommand : CommandBase
    {
        public override int CommandId => 0x0300;

        private EnsureOnlySelectedProjReferencesAreLoadedCommand(Package package) : base(package)
        {
            AddCommandToMenu(this.MenuItemCallback);
        }

        public static EnsureOnlySelectedProjReferencesAreLoadedCommand Instance
        {
            get;
            private set;
        }

        public static void Initialize(Package package)
        {
            Instance = new EnsureOnlySelectedProjReferencesAreLoadedCommand(package);
        }

        public void MenuItemCallback(object sender, EventArgs e)
        {
            if (ShowErrorMessageAndReturnTrueIfNoSolutionOpen())
            {
                return;
            }

            DteUtil.EnsureSelectedProjReferencesAreLoadedCommand(ServiceProvider, true);
        }
    }
}
