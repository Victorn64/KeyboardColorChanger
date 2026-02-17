using System;
using System.Windows.Forms;

namespace keybord_color_changer
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Запускаем форму, но не через Application.Run(form), 
            // а просто создаем экземпляр. Логика трея внутри Form1.
            Form1 mainForm = new Form1();
            
            Application.Run(); // Запускаем цикл сообщений без привязки к форме
        }
    }
}
