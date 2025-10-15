using ImageMagick;
using OfficeOpenXml;

namespace EmailPDFMatchKeyword
{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Set EPPlus license context before any EPPlus usage
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ApplicationConfiguration.Initialize();
            InitializeMagick();
            Application.Run(new MainForm());

        }
        private static void InitializeMagick()
        {
          string ghostscriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ghostscript", "bin");
          if (Directory.Exists(ghostscriptPath))
            MagickNET.SetGhostscriptDirectory(ghostscriptPath);

          string magickTempPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MagickTemp");
          if (!Directory.Exists(magickTempPath))
            Directory.CreateDirectory(magickTempPath);

          MagickNET.SetTempDirectory(magickTempPath);
        }
    } 
}