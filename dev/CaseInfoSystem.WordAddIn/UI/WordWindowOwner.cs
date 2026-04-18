using System;
using System.Windows.Forms;

namespace CaseInfoSystem.WordAddIn.UI
{
    internal sealed class WordWindowOwner : IWin32Window
    {
        public WordWindowOwner(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; private set; }
    }
}
