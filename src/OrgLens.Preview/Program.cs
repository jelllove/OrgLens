using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OrgLens.Desktop;

namespace OrgLens.Preview
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length == 1 && args[0] == "--smoke-test")
            {
                try
                {
                    SmokeTests.Run();
                    Console.WriteLine("OrgLens UI smoke checks passed. No Outlook data was accessed.");
                    return 0;
                }
                catch (InvalidOperationException exception)
                {
                    Console.Error.WriteLine("OrgLens UI smoke check failed: " + exception.Message);
                    return 1;
                }
            }
            if (args.Length == 2 && args[0] == "--screenshot")
            {
                try
                {
                    string path = args[1];
                    if (!Path.IsPathRooted(path) || !string.Equals(Path.GetFullPath(path), path, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.Error.WriteLine("--screenshot requires a fully qualified absolute file path.");
                        return 2;
                    }
                    return CaptureScreenshot(path);
                }
                catch (ArgumentException exception)
                {
                    Console.Error.WriteLine(exception.Message);
                    return 2;
                }
                catch (NotSupportedException exception)
                {
                    Console.Error.WriteLine(exception.Message);
                    return 2;
                }
                catch (IOException exception)
                {
                    Console.Error.WriteLine(exception.Message);
                    return 2;
                }
            }
            if (args.Length != 0)
            {
                Console.Error.WriteLine("Usage: OrgLens.Preview.exe [--smoke-test | --screenshot <absolute-png-path>]");
                return 2;
            }
            using (var form = new SettingsForm(new DemoService()))
            {
                Application.Run(form);
            }
            return 0;
        }

        private static int CaptureScreenshot(string path)
        {
            int result = 0;
            using (var form = new SettingsForm(new DemoService()))
            {
                form.Shown += (sender, args) => form.BeginInvoke((MethodInvoker)(() =>
                {
                    try
                    {
                        form.Update();
                        using (var bitmap = new Bitmap(form.Width, form.Height))
                        {
                            form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                            bitmap.Save(path, ImageFormat.Png);
                        }
                        Console.WriteLine("Sample-only screenshot saved: " + path);
                    }
                    catch (IOException exception)
                    {
                        result = ScreenshotFailed(exception.Message);
                    }
                    catch (UnauthorizedAccessException exception)
                    {
                        result = ScreenshotFailed(exception.Message);
                    }
                    catch (ExternalException exception)
                    {
                        result = ScreenshotFailed(exception.Message);
                    }
                    catch (ArgumentException exception)
                    {
                        result = ScreenshotFailed(exception.Message);
                    }
                    finally
                    {
                        form.Close();
                    }
                }));
                Application.Run(form);
            }
            return result;
        }

        private static int ScreenshotFailed(string message)
        {
            Console.Error.WriteLine("Could not save sample screenshot: " + message);
            return 2;
        }
    }
}
