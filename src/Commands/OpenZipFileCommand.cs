using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace OpenZipProject
{
    internal sealed class OpenZipFileCommand
    {
        private static DTE2 _dte;

        public static async Task InitializeAsync(AsyncPackage package, DTE2 dte)
        {
            _dte = dte;
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Assumes.Present(commandService);

            var menuCommandID = new CommandID(PackageGuids.guidOpenZipProjectCmd, PackageIds.OpenZipFileCommandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID)
            {
                ParametersDescription = "*"
            };

            commandService.AddCommand(menuItem);
        }

        private static void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (e is OleMenuCmdEventArgs cmdArgs && cmdArgs.InValue is string arg && !string.IsNullOrEmpty(arg))
            {
                var url = arg.Replace("vszip://", "https://").Trim('\"');
                var tempFile = Path.GetTempFileName();

                using (var client = new WebClient())
                {
                    client.DownloadFile(url, tempFile);
                }

                var dialog = new FolderBrowserDialog();
                DialogResult result = dialog.ShowDialog();

                if (result == DialogResult.OK)
                {
                    var dirName = Path.GetFileNameWithoutExtension(url);
                    var targetDir = Path.Combine(dialog.SelectedPath, dirName);

                    PackageUtilities.EnsureOutputPath(targetDir);
                    ZipFile.ExtractToDirectory(tempFile, targetDir);

                    var solution = Directory.EnumerateFiles(targetDir, "*.sln", SearchOption.AllDirectories).FirstOrDefault();

                    if (!string.IsNullOrEmpty(solution))
                    {
                        _dte.Solution.Open(solution);
                    }
                }
            }
        }
    }
}
