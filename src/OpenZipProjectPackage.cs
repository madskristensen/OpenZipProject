using System.Diagnostics;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using Microsoft;
using EnvDTE;
using EnvDTE80;

namespace OpenZipProject
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid("d2822e1b-2c7a-44b2-9be3-50a0428e5995")]
    [ProvideAppCommandLine("zip", typeof(OpenZipProjectPackage), Arguments = "1", DemandLoad = 1)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class OpenZipProjectPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var cmdline = await GetServiceAsync(typeof(SVsAppCommandLine)) as IVsAppCommandLine;
            Assumes.Present(cmdline);

            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            Assumes.Present(dte);

            await OpenZipFileCommand.InitializeAsync(this, dte);

            ErrorHandler.ThrowOnFailure(cmdline.GetOption("zip", out var isPresent, out var arg));

            if (isPresent == 1)
            {
                Command cmd = dte.Commands.Item("File.OpenFromZipFile");

                if (cmd != null && cmd.IsAvailable)
                {
                    dte.ExecuteCommand(cmd.Name, arg);
                }
            }
        }
    }
}
