// Mastercam 2024 NET-Script prototype.
// Select one or more posted NC files and send them to CNC 智能程序单.

var app = new App();
app.Run();

public class App
{
    public void Run()
    {
        var appPath = ResolveApplicationPath();
        if (System.String.IsNullOrWhiteSpace(appPath))
        {
            return;
        }

        using (var dialog = new System.Windows.Forms.OpenFileDialog())
        {
            dialog.Title = "选择已后处理的 NC 程序";
            dialog.Filter = "NC 程序 (*.NC;*.nc)|*.NC;*.nc|所有文件|*.*";
            dialog.Multiselect = true;
            dialog.CheckFileExists = true;
            dialog.RestoreDirectory = true;

            if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            var arguments = "--import-nc " + System.String.Join(
                " ",
                System.Array.ConvertAll(dialog.FileNames, QuoteArgument));

            try
            {
                var startInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = appPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    WorkingDirectory = System.IO.Path.GetDirectoryName(appPath)
                };
                System.Diagnostics.Process.Start(startInfo);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "启动程序单失败：\n" + ex.Message,
                    "Mastercam 程序单连接失败",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }

    private static string ResolveApplicationPath()
    {
        var appPath = System.Environment.GetEnvironmentVariable("CNC_PROGRAM_SHEET_EXE");
        if (!System.String.IsNullOrWhiteSpace(appPath) && System.IO.File.Exists(appPath))
        {
            return appPath;
        }

        var defaultPath = System.IO.Path.Combine(
            System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles),
            "CNC Program Sheet",
            "CNCProgramSheet.exe");
        if (System.IO.File.Exists(defaultPath))
        {
            return defaultPath;
        }

        var appData = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData);
        var settingsFolder = System.IO.Path.Combine(appData, "CNCProgramSheet");
        var settingsPath = System.IO.Path.Combine(settingsFolder, "mastercam_bridge_path.txt");
        try
        {
            if (System.IO.File.Exists(settingsPath))
            {
                var rememberedPath = System.IO.File.ReadAllText(settingsPath).Trim();
                if (System.IO.File.Exists(rememberedPath))
                {
                    return rememberedPath;
                }
            }
        }
        catch
        {
            // Continue to the one-time file picker if the saved location is stale.
        }

        using (var picker = new System.Windows.Forms.OpenFileDialog())
        {
            picker.Title = "首次使用：选择 CNC 智能程序单程序";
            picker.Filter = "CNC 智能程序单 (CNCProgramSheet.exe)|CNCProgramSheet.exe|应用程序 (*.exe)|*.exe";
            picker.FileName = "CNCProgramSheet.exe";
            picker.CheckFileExists = true;
            if (picker.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return "";
            }

            try
            {
                System.IO.Directory.CreateDirectory(settingsFolder);
                System.IO.File.WriteAllText(settingsPath, picker.FileName);
            }
            catch
            {
                // The selected path is still usable for this invocation.
            }
            return picker.FileName;
        }
    }

    private static string QuoteArgument(string value)
    {
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
