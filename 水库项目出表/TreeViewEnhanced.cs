using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 水库项目出表
{
    public class TreeViewEnhanced : TreeView
    {
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x203)
            {
                m.Result = IntPtr.Zero;
            }
            else base.WndProc(ref m);
        }
    }
}
