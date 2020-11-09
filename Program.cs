using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoSendMail
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (MessageBox.Show("このソフトを利用するにあたり\n誤送信などの問題が発生しても作成側は責任を持ちません。\nすべて自己責任での使用をお願い致します。", "利用規約", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Application.SetHighDpiMode(HighDpiMode.SystemAware);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new AutoSendMail());
            }
            else
            {
                Application.Exit();
            }
        }
    }
}
