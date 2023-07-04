using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace VSIXProject1
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(VSIXProject1Package.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class VSIXProject1Package : AsyncPackage
    {
        /// <summary>
        /// VSIXProject1Package GUID string.
        /// </summary>
        public const string PackageGuidString = "a4308b57-1142-42cf-b4bf-750f2a799c8e";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>

        private CommandEvents commandEvents;
        private DTE2 dte;
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(DisposalToken);
            dte = await GetServiceAsync(typeof(DTE)) as DTE2;           
            commandEvents = dte.Events.CommandEvents[typeof(VSConstants.VSStd97CmdID).GUID.ToString("B"), (int)VSConstants.VSStd97CmdID.F1Help];
            commandEvents.AfterExecute += OnAfterExecute;                        
        }

        private void OnAfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
        {            
            var activeWindow = dte.ActiveWindow;

            var contextAttributes = activeWindow.DTE.ContextAttributes;
            contextAttributes.Refresh();

            var attributes = new List<string>();
            try
            {
                ContextAttributes highPri = contextAttributes?.HighPriorityAttributes;
                highPri.Refresh();
                if (highPri != null)
                {
                    foreach (ContextAttribute CA in highPri)
                    {
                        var values = new List<string>();
                        foreach (string value in (ICollection)CA.Values)
                        {
                            values.Add(value);
                        }
                        var attribute = CA.Name + "=" + string.Join(";", values.ToArray());
                        attributes.Add(CA.Name + "=");
                    }
                }
            }
            catch
            {                
                // ignore this exception-- means there's no High Pri values here
            }             

            // fetch context attributes that are not high-priority
            foreach (ContextAttribute CA in contextAttributes)
            {
                var values = new List<string>();
                foreach (string value in (ICollection)CA.Values)
                {
                    values.Add(value);
                }
                string attribute = CA.Name + "=" + string.Join(";", values.ToArray());
                attributes.Add(attribute);
            }

            HandleHelpCall(attributes);
        }

        private void HandleHelpCall(List<string> attributes)
        {
            var keywords = attributes.SingleOrDefault(a => a.StartsWith("keyword"));
            var keyword = keywords.Split(';').First().Split('=')[1];
            System.Diagnostics.Process.Start($"http://google.com/search?q={keyword}");
        }
    }

    #endregion
}
