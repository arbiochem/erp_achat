using System;
using System.Windows.Forms;

namespace arbioApp.Utils
{
    public static class EventHandlers
    {
        public static void KeyDownEnterHandler(object sender, KeyEventArgs e, EventHandler clickHandler)
        {
            if (e.KeyCode == Keys.Enter)
            {
                clickHandler(sender, e);
            }
        }
    }
}
