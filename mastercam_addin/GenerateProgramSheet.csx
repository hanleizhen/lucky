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

        var selectedFiles = new string[0];
        var initialFolder = ResolvePostNcFolder();
        if (ReadSequenceSearchEnabled())
        {
            selectedFiles = ChooseProgramSequenceGroup(
                initialFolder, ReadSequenceSearchCount());
            if (selectedFiles == null)
            {
                return;
            }
        }

        if (selectedFiles.Length == 0)
        {
            using (var dialog = new System.Windows.Forms.OpenFileDialog())
            {
                dialog.Title = mode == "NORMAL"
                    ? "选择已后处理的 NC 程序并填写程序单"
                    : "选择已后处理、要执行 " + mode + " 修改的 NC 程序";
                dialog.Filter = "NC 程序 (*.NC;*.nc)|*.NC;*.nc|所有文件|*.*";
                dialog.Multiselect = true;
                dialog.CheckFileExists = true;
                dialog.RestoreDirectory = true;
                dialog.InitialDirectory = initialFolder;

                if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    return;
                }
                selectedFiles = dialog.FileNames;
            }
        }

        var arguments = "--from-mastercam ";
        if (mode != "NORMAL")
        {
            arguments += "--conversion-mode " + mode + " --apply-conversion ";
        }
        arguments += "--import-nc " + System.String.Join(
            " ", System.Array.ConvertAll(selectedFiles, QuoteArgument));

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

    private static bool ReadSequenceSearchEnabled()
    {
        try
        {
            var path = System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "CNCProgramSheet", "mastercam_sequence_search_enabled.txt");
            bool enabled;
            if (System.IO.File.Exists(path) &&
                bool.TryParse(System.IO.File.ReadAllText(path).Trim(), out enabled))
            {
                return enabled;
            }
        }
        catch
        {
            // Use the first-run default if the optional preference is unreadable.
        }
        return true;
    }

    private static int ReadSequenceSearchCount()
    {
        try
        {
            var path = System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "CNCProgramSheet", "mastercam_sequence_search_count.txt");
            int count;
            if (System.IO.File.Exists(path) &&
                int.TryParse(System.IO.File.ReadAllText(path).Trim(), out count))
            {
                return System.Math.Max(1, System.Math.Min(100, count));
            }
        }
        catch
        {
            // Use the default group size if the optional preference is unreadable.
        }
        return 11;
    }

    private static string[] ChooseProgramSequenceGroup(string folder, int defaultCount)
    {
        if (!System.IO.Directory.Exists(folder))
        {
            return new string[0];
        }

        using (var form = new System.Windows.Forms.Form())
        {
            form.Text = "按程序名查找连续 NC 程序组";
            form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.ClientSize = new System.Drawing.Size(620, 430);

            var nameLabel = new System.Windows.Forms.Label();
            nameLabel.Text = "输入起始程序名（例如 500.NC；只匹配相同前缀）：";
            nameLabel.AutoSize = true;
            nameLabel.Left = 18;
            nameLabel.Top = 18;
            form.Controls.Add(nameLabel);

            var nameInput = new System.Windows.Forms.TextBox();
            nameInput.Left = 20;
            nameInput.Top = 44;
            nameInput.Width = 390;
            form.Controls.Add(nameInput);

            var countLabel = new System.Windows.Forms.Label();
            countLabel.Text = "连续数量：";
            countLabel.AutoSize = true;
            countLabel.Left = 425;
            countLabel.Top = 48;
            form.Controls.Add(countLabel);

            var countInput = new System.Windows.Forms.NumericUpDown();
            countInput.Left = 500;
            countInput.Top = 44;
            countInput.Width = 95;
            countInput.Minimum = 1;
            countInput.Maximum = 100;
            countInput.Value = System.Math.Max(1, System.Math.Min(100, defaultCount));
            form.Controls.Add(countInput);

            var results = new System.Windows.Forms.ListBox();
            results.Left = 20;
            results.Top = 82;
            results.Width = 575;
            results.Height = 270;
            form.Controls.Add(results);

            var search = new System.Windows.Forms.Button();
            search.Text = "查找连续程序";
            search.Left = 20;
            search.Top = 370;
            search.Width = 130;
            search.Height = 34;
            form.Controls.Add(search);

            var manual = new System.Windows.Forms.Button();
            manual.Text = "手动多选文件";
            manual.Left = 165;
            manual.Top = 370;
            manual.Width = 130;
            manual.Height = 34;
            form.Controls.Add(manual);

            var useGroup = new System.Windows.Forms.Button();
            useGroup.Text = "处理这组程序";
            useGroup.Left = 350;
            useGroup.Top = 370;
            useGroup.Width = 130;
            useGroup.Height = 34;
            useGroup.Enabled = false;
            form.Controls.Add(useGroup);

            var cancel = new System.Windows.Forms.Button();
            cancel.Text = "取消";
            cancel.Left = 490;
            cancel.Top = 370;
            cancel.Width = 105;
            cancel.Height = 34;
            cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            form.Controls.Add(cancel);
            form.CancelButton = cancel;

            string[] foundFiles = new string[0];
            string[] selectedFiles = null;
            System.Action invalidateResults = () =>
            {
                foundFiles = new string[0];
                results.Items.Clear();
                useGroup.Enabled = false;
            };
            nameInput.TextChanged += (sender, args) => invalidateResults();
            countInput.ValueChanged += (sender, args) => invalidateResults();
            search.Click += (sender, args) =>
            {
                foundFiles = FindSequentialProgramFiles(
                    folder, nameInput.Text, System.Decimal.ToInt32(countInput.Value));
                results.Items.Clear();
                foreach (var file in foundFiles)
                {
                    results.Items.Add(System.IO.Path.GetFileName(file));
                }
                useGroup.Enabled = foundFiles.Length > 0;
                if (foundFiles.Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        form,
                        "没有找到以该文件名开始的连续 NC 程序。请确认文件名、数量和 NC 文件夹。",
                        "未找到连续程序",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            };
            useGroup.Click += (sender, args) =>
            {
                selectedFiles = foundFiles;
                form.DialogResult = System.Windows.Forms.DialogResult.OK;
                form.Close();
            };
            manual.Click += (sender, args) =>
            {
                selectedFiles = new string[0];
                form.DialogResult = System.Windows.Forms.DialogResult.OK;
                form.Close();
            };

            var result = form.ShowDialog();
            return result == System.Windows.Forms.DialogResult.OK ? selectedFiles : null;
        }
    }

    private static string[] FindSequentialProgramFiles(string folder, string inputName, int count)
    {
        if (System.String.IsNullOrWhiteSpace(inputName) || !System.IO.Directory.Exists(folder))
        {
            return new string[0];
        }
        var fileName = System.IO.Path.GetFileName(inputName.Trim().Trim('"'));
        var extension = System.IO.Path.GetExtension(fileName);
        if (System.String.IsNullOrEmpty(extension))
        {
            extension = ".NC";
            fileName += extension;
        }
        if (!extension.Equals(".NC", System.StringComparison.OrdinalIgnoreCase))
        {
            return new string[0];
        }
        var stem = System.IO.Path.GetFileNameWithoutExtension(fileName);
        var match = System.Text.RegularExpressions.Regex.Match(stem, @"^(.*?)(\d+)$");
        if (!match.Success)
        {
            return new string[0];
        }
        int firstNumber;
        if (!int.TryParse(match.Groups[2].Value, out firstNumber))
        {
            return new string[0];
        }

        var prefix = match.Groups[1].Value;
        var numberWidth = match.Groups[2].Value.Length;
        var files = new System.Collections.Generic.List<string>();
        var maximum = System.Math.Max(1, System.Math.Min(100, count));
        for (int offset = 0; offset < maximum; offset++)
        {
            if (firstNumber > int.MaxValue - offset)
            {
                break;
            }
            var number = (firstNumber + offset).ToString("D" + numberWidth);
            var candidate = System.IO.Path.Combine(folder, prefix + number + extension);
            if (!System.IO.File.Exists(candidate))
            {
                break;
            }
            files.Add(candidate);
        }
        return files.ToArray();
    }

    private static string ResolvePostNcFolder()
    {
        try
        {
            var localAppData = System.Environment.GetFolderPath(
                System.Environment.SpecialFolder.LocalApplicationData);
            var configuredFolderFile = System.IO.Path.Combine(
                localAppData, "CNCProgramSheet", "mastercam_nc_folder.txt");
            if (System.IO.File.Exists(configuredFolderFile))
            {
                var configuredFolder = System.IO.File.ReadAllText(configuredFolderFile).Trim();
                if (System.IO.Directory.Exists(configuredFolder))
                {
                    return configuredFolder;
                }
            }

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
