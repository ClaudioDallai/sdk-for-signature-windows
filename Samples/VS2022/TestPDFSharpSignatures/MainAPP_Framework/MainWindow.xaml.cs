using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MainAPP_Framework
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string relativePath = @"..\..\..\TestPDFSharpSignatures\bin\Debug\net8.0-windows\TestPDFSharpSignatures.exe";
            string SignModuleExePath = System.IO.Path.GetFullPath(System.IO.Path.Combine(basePath, relativePath));

            string[] args = new string[]
            {
                "Sign Reason Crazy Placeholder",
                "Italia, Napoli",
                "Tommasino",
                "Verdi"
            };

            string arguments = string.Join(" ", args.Select(a => $"\"{a}\""));

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = SignModuleExePath,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            Process.Start(psi);
        }
    }
}
