// Mastercam 2021–2027 NET-Script bridge for CNC 智能程序单.

var app = new App();
app.Run();

public class App
{
    public void Run()
    {
        var mode = ChooseMode();
        if (System.String.IsNullOrWhiteSpace(mode))
        {
            return;
        }

        var appPath = ResolveApplicationPath();
        if (System.String.IsNullOrWhiteSpace(appPath))
        {
            return;
        }

        using (var dialog = new System.Windows.Forms.OpenFileDialog())
        {
            dialog.Title = mode == "NORMAL"
                ? "选择已后处理的 NC 程序并填写程序单"
                : "选择已后处理、要执行 " + mode + " 修改的 NC 程序";
            dialog.Filter = "NC 程序 (*.NC;*.nc)|*.NC;*.nc|所有文件|*.*";
            dialog.Multiselect = true;
            dialog.CheckFileExists = true;
            dialog.RestoreDirectory = true;
            dialog.InitialDirectory = ResolvePostNcFolder();

            if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            var arguments = "--from-mastercam ";
            if (mode != "NORMAL")
            {
                arguments += "--conversion-mode " + mode + " --apply-conversion ";
            }
            arguments += "--import-nc " + System.String.Join(
                " ", System.Array.ConvertAll(dialog.FileNames, QuoteArgument));

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

    private static string ResolvePostNcFolder()
    {
        try
        {
            var documents = System.Environment.GetFolderPath(
                System.Environment.SpecialFolder.MyDocuments);
            var processPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
            var installDirectory = new System.IO.DirectoryInfo(
                System.IO.Path.GetDirectoryName(processPath));
            var year = "";
            while (installDirectory != null)
            {
                var match = System.Text.RegularExpressions.Regex.Match(
                    installDirectory.Name, @"20\d{2}");
                if (match.Success)
                {
                    year = match.Value;
                    break;
                }
                installDirectory = installDirectory.Parent;
            }
            if (!System.String.IsNullOrEmpty(year))
            {
                var profile = System.IO.Path.Combine(
                    documents, "My Mastercam " + year, "Mastercam");
                var candidates = new string[] {
                    System.IO.Path.Combine(profile, "Mill", "NC"),
                    System.IO.Path.Combine(profile, "NC"),
                    profile
                };
                foreach (var candidate in candidates)
                {
                    if (System.IO.Directory.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }
        }
        catch
        {
            // Fall back to the file dialog's normal recent-folder behavior.
        }
        return System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
    }

    private static string ChooseMode()
    {
        using (var form = new System.Windows.Forms.Form())
        {
            form.Text = "CNC 程序单 - 选择处理方式";
            form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.ClientSize = new System.Drawing.Size(420, 235);

            var label = new System.Windows.Forms.Label();
            label.Text = "请选择 NC 程序处理方式：";
            label.AutoSize = true;
            label.Left = 22;
            label.Top = 20;
            form.Controls.Add(label);

            string selected = "";
            string[] modes = { "M06", "M304", "铜工", "正常" };
            for (int i = 0; i < modes.Length; i++)
            {
                var button = new System.Windows.Forms.Button();
                button.Text = modes[i];
                button.Width = 82;
                button.Height = 38;
                button.Left = 20 + i * 96;
                button.Top = 68;
                var mode = modes[i];
                button.Click += (sender, args) =>
                {
                    selected = mode == "铜工" ? "COPPER" : (mode == "正常" ? "NORMAL" : mode);
                    form.DialogResult = System.Windows.Forms.DialogResult.OK;
                    form.Close();
                };
                form.Controls.Add(button);
            }

            var cancel = new System.Windows.Forms.Button();
            cancel.Text = "取消";
            cancel.Width = 82;
            cancel.Height = 30;
            cancel.Left = 318;
            cancel.Top = 180;
            cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            form.Controls.Add(cancel);
            form.CancelButton = cancel;

            return form.ShowDialog() == System.Windows.Forms.DialogResult.OK ? selected : "";
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
