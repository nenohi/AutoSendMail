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
            if (MessageBox.Show("���̃\�t�g�𗘗p����ɂ�����\n�둗�M�Ȃǂ̖�肪�������Ă��쐬���͐ӔC�������܂���B\n���ׂĎ��ȐӔC�ł̎g�p�����肢�v���܂��B", "���p�K��", MessageBoxButtons.YesNo) == DialogResult.Yes)
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
