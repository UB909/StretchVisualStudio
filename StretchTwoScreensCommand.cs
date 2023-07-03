using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace StretchVisualstudio {
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class StretchTwoScreensCommand {
    /// <summary>
    /// Command ID.
    /// </summary>
    public const int CommandId = 0x0100;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("cf3c751e-1022-48b3-8d25-8c7763cce0e1");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly AsyncPackage package;

    /// <summary>
    /// Initializes a new instance of the <see cref="StretchTwoScreensCommand"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    /// <param name="commandService">Command service to add command to, not null.</param>
    private StretchTwoScreensCommand(AsyncPackage package, OleMenuCommandService commandService) {
      this.package = package ?? throw new ArgumentNullException(nameof(package));
      commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

      var menuCommandID = new CommandID(CommandSet, CommandId);
      var menuItem = new MenuCommand(this.Execute, menuCommandID);
      commandService.AddCommand(menuItem);
    }

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static StretchTwoScreensCommand Instance {
      get;
      private set;
    }

    /// <summary>
    /// Gets the service provider from the owner package.
    /// </summary>
    private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider {
      get {
        return this.package;
      }
    }

    /// <summary>
    /// Initializes the singleton instance of the command.
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    public static async Task InitializeAsync(AsyncPackage package) {
      // Switch to the main thread - the call to AddCommand in StretchTwoScreensCommand's constructor requires
      // the UI thread.
      await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

      OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
      Instance = new StretchTwoScreensCommand(package, commandService);
    }

    private const int ABM_GETTASKBARPOS = 5;

    [System.Runtime.InteropServices.DllImport("shell32.dll")]
    private static extern IntPtr SHAppBarMessage(int msg, ref APPBARDATA data);

    private struct APPBARDATA {
      public int cbSize;
      public IntPtr hWnd;
      public int uCallbackMessage;
      public int uEdge;
      public RECT rc;
      public IntPtr lParam;
    }

    private struct RECT {
      public int left, top, right, bottom;
    }

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    const uint SWP_NOZORDER = 0x0004;


    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    /// <summary>
    /// This function is the callback used to execute the command when the menu item is clicked.
    /// See the constructor to see how the menu item is associated with this function using
    /// OleMenuCommandService service and MenuCommand class.
    /// </summary>
    /// <param name="sender">Event sender.</param>
    /// <param name="e">Event args.</param>
    private void Execute(object sender, EventArgs e) {
      ThreadHelper.ThrowIfNotOnUIThread();
      Process[] processlist = Process.GetProcesses();

      // get size of taskbar
      APPBARDATA data = new APPBARDATA();
      data.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(data);
      SHAppBarMessage(ABM_GETTASKBARPOS, ref data);

      List<string> l = new List<string>();

      //foreach (Process process in processlist) {
      Process process = Process.GetCurrentProcess();
        if (!String.IsNullOrEmpty(process.MainWindowTitle)) {
          if (((process.ProcessName.ToLower().Contains("devenv")) && (process.MainWindowTitle.ToLower().Contains("microsoft visual studio"))) ||
              ((process.ProcessName.ToLower().Contains("code")) && (process.MainWindowTitle.ToLower().Contains("visual studio code")))) {
          ShowWindow(process.MainWindowHandle, 1);// Show Normal (not maximized)
          SetWindowPos(
                process.MainWindowHandle,
                IntPtr.Zero,
                (int)System.Windows.SystemParameters.VirtualScreenLeft,
                (int)System.Windows.SystemParameters.VirtualScreenTop,
                (int)System.Windows.SystemParameters.VirtualScreenWidth,
                (int)System.Windows.SystemParameters.VirtualScreenHeight - (data.rc.bottom - data.rc.top),
                SWP_NOZORDER);
          }
          l.Add(process.ProcessName + " - " + process.MainWindowTitle);
        }
      //}


      //string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
      //string title = "StretchTwoScreensCommand";

      //// Show a message box to prove we were here
      //VsShellUtilities.ShowMessageBox(
      //    this.package,
      //    message,
      //    title,
      //    OLEMSGICON.OLEMSGICON_INFO,
      //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
      //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
    }
  }
}
