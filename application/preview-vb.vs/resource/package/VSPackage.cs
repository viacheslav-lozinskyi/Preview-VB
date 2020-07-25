
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace resource.package
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(CONSTANT.NAME, CONSTANT.DESCRIPTION, CONSTANT.VERSION)]
    [Guid(CONSTANT.GUID)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.FirstLaunchSetup_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class PreviewVBPackage : AsyncPackage
    {
        internal static class CONSTANT
        {
            public const string COPYRIGHT = "Copyright (c) 2020 by Viacheslav Lozinskyi. All rights reserved.";
            public const string DESCRIPTION = "Quick preview of VB files";
            public const string EXTENSION = ".VB";
            public const string GUID = "E91ECA80-9650-47F4-91CD-EFBA501DDAA8";
            public const string NAME = "Preview-VB";
            public const string VERSION = "1.0.2";
        }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            {
                cartridge.AnyPreview.Connect();
                cartridge.AnyPreview.Register(cartridge.AnyPreview.MODE.PREVIEW, CONSTANT.EXTENSION, new preview.VB());
            }
        }

        protected override int QueryClose(out bool canClose)
        {
            {
                cartridge.AnyPreview.Disconnect();
                canClose = true;
            }
            return VSConstants.S_OK;
        }
    }
}
