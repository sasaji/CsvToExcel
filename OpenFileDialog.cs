using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace CsvToExcel
{
    /// <summary>
    /// The extensible OpenFileDialog
    /// </summary>
    public class OpenFileDialog : IDisposable
	{
		// The maximum number of characters permitted in a path
		private const int _MAX_PATH = 260;

		// The "control ID" of the content window inside the OpenFileDialog
		// See the accompanying article to learn how I discovered it
		private const int _CONTENT_PANEL_ID = 0x0461;

		// A constant that determines the spacing between panels inside the OpenFileDialog
		private const int _PANEL_GAP_FACTOR = 3;

		/// <summary>
		/// Clients can implement handlers of this type to catch "selection changed" events
		/// </summary>
		public delegate void SelectionChangedHandler( string path );

		/// <summary>
		/// This event is fired whenever the user selects an item in the dialog
		/// </summary>
		//public event SelectionChangedHandler SelectionChanged;

		// unmanaged memory buffers to hold the file name (with and without full path)
		private IntPtr _fileNameBuffer;
		private IntPtr _fileTitleBuffer;

        // title
        private IntPtr _titleBuffer;

		// the OPENFILENAME structure, used to control the appearance and behaviour of the OpenFileDialog
		private OpenFileName _ofn;

		// user-supplied control that gets placed inside the OpenFileDialog
		private System.Windows.Forms.Control _userControl;

		// unmanaged memory buffer that holds the Win32 dialog template
		private IntPtr _ipTemplate;

		/// <summary>
		/// Sets up the data structures necessary to display the OpenFileDialog
		/// </summary>
		/// <param name="defaultExtension">The file extension to use if the user doesn't specify one (no "." required)</param>
		/// <param name="fileName">You can specify a filename to appear in the dialog, although the user can change it</param>
		/// <param name="filter">See the documentation for the OPENFILENAME structure for a description of filter strings</param>
		/// <param name="userPanel">Any Windows Forms control, it will be placed inside the OpenFileDialog</param>
		public OpenFileDialog( string defaultExtension, string fileName, string filter, System.Windows.Forms.Control userControl, string title )
		{
			// Need two buffers in unmanaged memory to hold the filename
			// Note: the multiplication by 2 is to allow for Unicode (16-bit) characters
			_fileNameBuffer = Marshal.AllocCoTaskMem( 2 * _MAX_PATH  );
			_fileTitleBuffer = Marshal.AllocCoTaskMem( 2 * _MAX_PATH );

			// Zero these two buffers
			byte[] zeroBuffer = new byte [2 * (_MAX_PATH+1)];
			for( int i = 0; i < 2 * (_MAX_PATH+1); i++ ) zeroBuffer[i] = 0;
			Marshal.Copy( zeroBuffer, 0, _fileNameBuffer, 2 * _MAX_PATH );
			Marshal.Copy( zeroBuffer, 0, _fileTitleBuffer, 2 * _MAX_PATH );

            _titleBuffer = Marshal.AllocCoTaskMem(256);
            Marshal.Copy(zeroBuffer, 0, _titleBuffer, 256);

            // Create an in-memory Win32 dialog template; this will be a "child" window inside the FileOpenDialog
            // We have no use for this child window, except that its presence allows us to capture events when
            // the user interacts with the FileOpenDialog
            _ipTemplate = BuildDialogTemplate();

			// Populate the OPENFILENAME structure
			// The flags specified are the minimal set to get the appearance and behaviour we need
			_ofn.lStructSize = Marshal.SizeOf( _ofn );
			_ofn.lpstrFile = _fileNameBuffer;
			_ofn.nMaxFile = _MAX_PATH;
			_ofn.lpstrDefExt = Marshal.StringToCoTaskMemUni( defaultExtension );
			_ofn.lpstrFileTitle = _fileTitleBuffer;
			_ofn.nMaxFileTitle = _MAX_PATH;
			_ofn.lpstrFilter = Marshal.StringToCoTaskMemUni( filter );
			_ofn.Flags = OpenFileNameFlags.EnableHook | OpenFileNameFlags.EnableTemplateHandle | OpenFileNameFlags.EnableSizing | OpenFileNameFlags.Explorer | OpenFileNameFlags.HideReadOnly;
			_ofn.hInstance = _ipTemplate;
			_ofn.lpfnHook = new OfnHookProc(MyHookProc);
            _ofn.lpstrTitle = Marshal.StringToCoTaskMemUni(title);
			
			// copy initial file name into unmanaged memory buffer
			UnicodeEncoding ue = new UnicodeEncoding();
			byte[] fileNameBytes = ue.GetBytes( fileName );
			Marshal.Copy( fileNameBytes, 0, _fileNameBuffer, fileNameBytes.Length );

			// keep a reference to the user-supplied control
			_userControl = userControl;
		}

		/// <summary>
		/// The finalizer will release the unmanaged memory, if I should forget to call Dispose
		/// </summary>
		~OpenFileDialog()
		{
			Dispose( false );
		}

		/// <summary>
		/// Display the OpenFileDialog and allow user interaction
		/// </summary>
		/// <returns>true if the user clicked OK, false if they clicked cancel (or close)</returns>
		public bool Show()
		{
			return NativeMethods.GetOpenFileName( ref _ofn );
		}

		/// <summary>
		/// Builds an in-memory Win32 dialog template.  See documentation for DLGTEMPLATE.
		/// </summary>
		/// <returns>a pointer to an unmanaged memory buffer containing the dialog template</returns>
		private IntPtr BuildDialogTemplate()
		{
			// We must place this child window inside the standard FileOpenDialog in order to get any
			// notifications sent to our hook procedure.  Also, this child window must contain at least
			// one control.  We make no direct use of the child window, or its control.

			// Set up the contents of the DLGTEMPLATE
			DlgTemplate template = new DlgTemplate();

			// Allocate some unmanaged memory for the template structure, and copy it in
			IntPtr ipTemplate = Marshal.AllocCoTaskMem( Marshal.SizeOf(template) );
			Marshal.StructureToPtr( template, ipTemplate, true );
			return ipTemplate;
		}

		/// <summary>
		/// The hook procedure for window messages generated by the FileOpenDialog
		/// </summary>
		/// <param name="hWnd">the handle of the window at which this message is targeted</param>
		/// <param name="msg">the message identifier</param>
		/// <param name="wParam">message-specific parameter data</param>
		/// <param name="lParam">mess-specific parameter data</param>
		/// <returns></returns>
		public IntPtr MyHookProc( IntPtr hWnd, UInt16 msg, Int32 wParam, Int32 lParam )
		{
			if (hWnd == IntPtr.Zero)
				return IntPtr.Zero;

            // Behaviour is dependant on the message received
            switch (msg) {
                // WM_INITDIALOG - at this point the OpenFileDialog exists, so we pull the user-supplied control
                // into the FileOpenDialog now, using the SetParent API.
                case WindowMessage.InitDialog:
                    IntPtr hWndParent = NativeMethods.GetParent(hWnd);
                    NativeMethods.SetParent(_userControl.Handle, hWndParent);
                    return IntPtr.Zero;

                // WM_SIZE - the OpenFileDialog has been resized, so we'll resize the content and user-supplied
                // panel to fit nicely
                case WindowMessage.Size:
                    FindAndResizePanels(hWnd);
                    return IntPtr.Zero;

                // WM_NOTIFY - we're only interested in the CDN_SELCHANGE notification message:
                // we grab the currently-selected filename and fire our event
                case WindowMessage.Notify:
                    /*
                IntPtr ipNotify = new IntPtr( lParam );
                OfNotify ofNot = (OfNotify)Marshal.PtrToStructure( ipNotify, typeof(OfNotify) );
                UInt16 code = ofNot.hdr.code;
                if( code == CommonDlgNotification.SelChange )
                {
                    // This is the first time we can rely on the presence of the content panel
                    // Resize the content and user-supplied panels to fit nicely
                    FindAndResizePanels( hWnd );

                    // get the newly-selected path
                    IntPtr hWndParent = NativeMethods.GetParent( hWnd );
                    StringBuilder pathBuffer = new StringBuilder(_MAX_PATH);
                    UInt32 ret = NativeMethods.SendMessage( hWndParent, CommonDlgMessage.GetFilePath, _MAX_PATH, pathBuffer );
                    string path = pathBuffer.ToString();

                    // copy the string into the path buffer
                    UnicodeEncoding ue = new UnicodeEncoding();
                    byte[] pathBytes = ue.GetBytes( path );
                    Marshal.Copy( pathBytes, 0, _fileNameBuffer, pathBytes.Length );

                    // fire selection-changed event
                    if( SelectionChanged != null ) SelectionChanged( path );
                }
                */
                    return IntPtr.Zero;
                // We're not interested in every possible message; just return a NULL for those we don't care about
                default:
                    return IntPtr.Zero;
            }
        }

		/// <summary>
		/// Layout the content of the OpenFileDialog, according to the overall size of the dialog
		/// </summary>
		/// <param name="hWnd">handle of window that received the WM_SIZE message</param>
		private void FindAndResizePanels( IntPtr hWnd )
		{
			// The FileOpenDialog is actually of the parent of the specified window
			IntPtr hWndParent = NativeMethods.GetParent( hWnd );

			// The "content" window is the one that displays the filenames, tiles, etc.
			// The _CONTENT_PANEL_ID is a magic number - see the accompanying text to learn
			// how I discovered it.
			IntPtr hWndContent = NativeMethods.GetDlgItem( hWndParent, _CONTENT_PANEL_ID );

			Rectangle rcClient = new Rectangle( 0, 0, 0, 0 );
			Rectangle rcContent = new Rectangle( 0, 0, 0, 0 );

			// Get client rectangle of dialog
			RECT rcTemp = new RECT();
			NativeMethods.GetClientRect( hWndParent, ref rcTemp );
			rcClient.X = rcTemp.left;
			rcClient.Y = rcTemp.top;
			rcClient.Width = rcTemp.right - rcTemp.left;
			rcClient.Height = rcTemp.bottom - rcTemp.top;

			// The content window may not be present when the dialog first appears
			if( hWndContent != IntPtr.Zero )
			{
				// Find the dimensions of the content panel
				RECT rc = new RECT();
				NativeMethods.GetWindowRect( hWndContent, ref rc );

				// Translate these dimensions into the dialog's coordinate system
				POINT topLeft;
				topLeft.X = rc.left;
				topLeft.Y = rc.top;
				NativeMethods.ScreenToClient( hWndParent, ref topLeft );
				POINT bottomRight;
				bottomRight.X = rc.right;
				bottomRight.Y = rc.bottom;
				NativeMethods.ScreenToClient( hWndParent, ref bottomRight );
				rcContent.X = topLeft.X;
				rcContent.Width = bottomRight.X - topLeft.X;
				rcContent.Y = topLeft.Y;
				rcContent.Height = bottomRight.Y - topLeft.Y;

				// Shrink content panel's width
				int width = rcClient.Right - rcContent.Left;
                int height = rcContent.Bottom - rcClient.Top;
				//rcContent.Width = (width/2) + _PANEL_GAP_FACTOR;
                rcContent.Width = (width - 300) + _PANEL_GAP_FACTOR; //sasa
                //rcContent.Height = (height - 130) + _PANEL_GAP_FACTOR; //sasa
                NativeMethods.MoveWindow( hWndContent, rcContent.Left, rcContent.Top, rcContent.Width, rcContent.Height, true );
			}

            // Position the user-supplied control alongside the content panel
            Rectangle rcUser = new Rectangle( rcContent.Right + (2 * _PANEL_GAP_FACTOR), rcContent.Top, rcClient.Right - rcContent.Right - (3 * _PANEL_GAP_FACTOR), rcContent.Bottom - rcContent.Top );
            NativeMethods.MoveWindow( _userControl.Handle, rcUser.X, rcUser.Y, rcUser.Width, rcUser.Height, true );
		}

		/// <summary>
		/// returns the path currently selected by the user inside the OpenFileDialog
		/// </summary>
		public string SelectedPath
		{
			get
			{
				return Marshal.PtrToStringUni( _fileNameBuffer );
			}
		}

		#region IDisposable Members

		public void Dispose()
		{
			Dispose( true );
		}

		/// <summary>
		/// Free any unamanged memory used by this instance of OpenFileDialog
		/// </summary>
		/// <param name="disposing">true if called by Dispose, false otherwise</param>
		public void Dispose( bool disposing )
		{
			if( disposing )
			{
				GC.SuppressFinalize( this );
			}

			Marshal.FreeCoTaskMem( _fileNameBuffer );
			Marshal.FreeCoTaskMem( _fileTitleBuffer );
			Marshal.FreeCoTaskMem( _ipTemplate );
		}

		#endregion
	}
}
