using System.Windows.Forms;

namespace arbioApp.Utils.Controls
{
    public static class AddControl
    {
        public static Form ToForm(Form form, UserControl control)
        {
            form.Controls.Add(control);
            control.Dock = DockStyle.Fill;
            form.Show();
            return form;
        }
    }
}
