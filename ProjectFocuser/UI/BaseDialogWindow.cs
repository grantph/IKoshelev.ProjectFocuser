using Microsoft.VisualStudio.PlatformUI;

namespace ProjectFocuser.UI
{
    //https://michaelscodingspot.com/2017/12/18/visual-studio-2017-extension-development-tutorial-part-4-show-a-popup-window/?subscribe=success#blog_subscription-2
    public class BaseDialogWindow : DialogWindow
    {

        public BaseDialogWindow()
        {
            this.HasMaximizeButton = true;
            this.HasMinimizeButton = true;
        }
    }
}
