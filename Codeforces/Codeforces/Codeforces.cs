using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Diagnostics;
using System.Windows.Forms;

namespace Codeforces
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Codeforces
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("28a7372c-20af-4a23-b8ad-dcd285e4f102");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Codeforces"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private Codeforces(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Codeforces Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new Codeforces(package);
        }

        private const string verbOpen = "Open";
        private IVsWebBrowsingService browserService = null;
        public ProcessStartInfo startInfo = null;

        /// <summary>
        /// Gets the initialized instance of ProcessStartInfo for a ShellExecute Open command.
        /// </summary>
        public ProcessStartInfo StartInfo
        {
            get
            {
                if (startInfo == null)
                {
                    startInfo = new ProcessStartInfo();
                    startInfo.UseShellExecute = true;
                    startInfo.Verb = verbOpen;
                }

                return startInfo;
            }
        }
        /// <summary>
        /// Launches the specified Url either in the internal VS browser or the
        /// user's default web browser.
        /// </summary>
        /// <param name="browserService">VS's browser service for interacting with the internal browser.</param>
        /// <param name="launchUrl">Url to launch.</param>
        /// <param name="useInternalBrowser">true to use the internal browser; false to use the default browser.</param>
        private void LaunchWebBrowser(IVsWebBrowsingService browserService, string launchUrl, bool useInternalBrowser)
        {
            try
            {
                if (useInternalBrowser == true)
                {
                    // if set to use internal browser, then navigate via the browser service.
                    IVsWindowFrame ppFrame;
                    // passing 0 to the NavigateFlags allows the browser service to reuse open instances
                    // of the internal browser.
                    browserService.Navigate(launchUrl, 0, out ppFrame);
                }
                else
                {
                    // if not, launch the user's default browser by starting a new one.
                    StartInfo.FileName = launchUrl;
                    Process.Start(StartInfo);
                }
            }
            catch
            {
                // if the process could not be started, show an error.
                MessageBox.Show("Cannot launch this url.", "Extension Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            //string message = "This is the codeforces message box.";
            //string title = "Codeforces";

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.ServiceProvider,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            browserService = ServiceProvider.GetService(typeof(SVsWebBrowsingService)) as IVsWebBrowsingService;
            LaunchWebBrowser(browserService, "http://codeforces.com/contests", true);
        }
    }
}
