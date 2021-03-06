using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CsvToExcel
{
	/// <summary>
	/// Defines the shape of hook procedures that can be called by the OpenFileDialog
	/// </summary>
	internal delegate IntPtr OfnHookProc( IntPtr hWnd, UInt16 msg, Int32 wParam, Int32 lParam );

	/// <summary>
	/// Values that can be placed in the OPENFILENAME structure, we don't use all of them
	/// </summary>
	internal class OpenFileNameFlags
	{
		public const Int32 ReadOnly =				0x00000001;
		public const Int32 OverWritePrompt =		0x00000002;
		public const Int32 HideReadOnly =			0x00000004;
		public const Int32 NoChangeDir =			0x00000008;
		public const Int32 ShowHelp =				0x00000010;
		public const Int32 EnableHook =				0x00000020;
		public const Int32 EnableTemplate =			0x00000040;
		public const Int32 EnableTemplateHandle =	0x00000080;
		public const Int32 NoValidate =				0x00000100;
		public const Int32 AllowMultiSelect =		0x00000200;
		public const Int32 ExtensionDifferent =		0x00000400;
		public const Int32 PathMustExist =			0x00000800;
		public const Int32 FileMustExist =			0x00001000;
		public const Int32 CreatePrompt =			0x00002000;
		public const Int32 ShareAware =				0x00004000;
		public const Int32 NoReadOnlyReturn =		0x00008000;
		public const Int32 NoTestFileCreate =		0x00010000;
		public const Int32 NoNetworkButton =		0x00020000;
		public const Int32 NoLongNames =			0x00040000;
		public const Int32 Explorer =				0x00080000;
		public const Int32 NoDereferenceLinks =		0x00100000;
		public const Int32 LongNames =				0x00200000;
		public const Int32 EnableIncludeNotify =	0x00400000;
		public const Int32 EnableSizing =			0x00800000;
		public const Int32 DontAddToRecent =		0x02000000;
		public const Int32 ForceShowHidden =		0x10000000;
	};

	/// <summary>
	/// Values that can be placed in the FlagsEx field of the OPENFILENAME structure
	/// </summary>
	internal class OpenFileNameFlagsEx
	{
		public const Int32 NoPlacesBar =			0x00000001;
	};

	/// <summary>
	/// A small subset of the window messages that can be sent to the OpenFileDialog
	/// These are just the ones that this implementation is interested in
	/// </summary>
	internal class WindowMessage
	{
		public const UInt16 InitDialog =	0x0110;
		public const UInt16 Size =			0x0005;
		public const UInt16 Notify =		0x004E;
	};

	/// <summary>
	/// The possible notification messages that can be generated by the OpenFileDialog
	/// We only look for CDN_SELCHANGE
	/// </summary>
	internal class CommonDlgNotification
	{
		private const UInt16 First =			unchecked((UInt16)((UInt16)0 - (UInt16)601));
		
		public const UInt16 InitDone =			(First - 0x0000);
		public const UInt16 SelChange =			(First - 0x0001);
		public const UInt16 FolderChange =		(First - 0x0002);
		public const UInt16 ShareViolation =	(First - 0x0003);
		public const UInt16 Help =				(First - 0x0004);
		public const UInt16 FileOk =			(First - 0x0005);
		public const UInt16 TypeChange =		(First - 0x0006);
		public const UInt16 IncludeItem =		(First - 0x0007);
	}

	/// <summary>
	/// Messages that can be send to the common dialogs
	/// We only use CDM_GETFILEPATH
	/// </summary>
	internal class CommonDlgMessage
	{
		private const UInt16 User =			0x0400;
		private const UInt16 First =		User + 100;
		
		public const UInt16 GetFilePath =	First + 0x0001;
	};

	/// <summary>
	/// See the documentation for OPENFILENAME
	/// </summary>
	internal struct OpenFileName
	{ 
		public Int32		lStructSize; 
		public IntPtr		hwndOwner; 
		public IntPtr		hInstance; 
		public IntPtr		lpstrFilter; 
		public IntPtr		lpstrCustomFilter; 
		public Int32		nMaxCustFilter; 
		public Int32		nFilterIndex; 
		public IntPtr		lpstrFile; 
		public Int32		nMaxFile; 
		public IntPtr		lpstrFileTitle; 
		public Int32		nMaxFileTitle; 
		public IntPtr		lpstrInitialDir; 
		public IntPtr		lpstrTitle; 
		public Int32		Flags; 
		public Int16		nFileOffset; 
		public Int16		nFileExtension; 
		public IntPtr		lpstrDefExt; 
		public Int32		lCustData; 
		public OfnHookProc	lpfnHook; 
		public IntPtr		lpTemplateName;
		public IntPtr		pvReserved;
		public Int32		dwReserved;
		public Int32		FlagsEx;
	};

	/// <summary>
	/// Part of the notification messages sent by the common dialogs
	/// </summary>
	[StructLayout(LayoutKind.Explicit)]
	internal struct NMHDR
	{
		[FieldOffset(0)]	public IntPtr	hWndFrom;
		[FieldOffset(4)]	public UInt16	idFrom;
		[FieldOffset(8)]	public UInt16	code;
	};

	/// <summary>
	/// Part of the notification messages sent by the common dialogs
	/// </summary>
	[StructLayout(LayoutKind.Explicit)]
	internal struct OfNotify
	{
		[FieldOffset(0)]	public NMHDR	hdr;
		[FieldOffset(12)]	public IntPtr	ipOfn;
		[FieldOffset(16)]	public IntPtr	ipFile;
	};

    /// <summary>
    /// Win32 window style constants
    /// We use them to set up our child window
    /// </summary>
    internal class DlgStyle
	{
		public const Int32 DsSetFont =		0x00000040;
		public const Int32 Ds3dLook =		0x00000004;
		public const Int32 DsControl =		0x00000400;
		public const Int32 WsChild =		0x40000000;
		public const Int32 WsClipSiblings =	0x04000000;
		public const Int32 WsVisible =		0x10000000;
		public const Int32 WsGroup =		0x00020000;
		public const Int32 SsNotify =		0x00000100;
	};

	/// <summary>
	/// Win32 "extended" window style constants
	/// </summary>
	internal class ExStyle
	{
		public const Int32 WsExNoParentNotify =	0x00000004;
		public const Int32 WsExControlParent =	0x00010000;
	};

	/// <summary>
	/// An in-memory Win32 dialog template
	/// Note: this has a very specific structure with a single static "label" control
	/// See documentation for DLGTEMPLATE and DLGITEMTEMPLATE
	/// </summary>
	[StructLayout(LayoutKind.Sequential)]
	internal class DlgTemplate
	{
		// The dialog template - see documentation for DLGTEMPLATE
		public Int32 style =			DlgStyle.Ds3dLook | DlgStyle.DsControl | DlgStyle.WsChild | DlgStyle.WsClipSiblings | DlgStyle.SsNotify;
		public Int32 extendedStyle =	ExStyle.WsExControlParent;
		public Int16 numItems =			1;
		public Int16 x =				0;
		public Int16 y =				0;
		public Int16 cx =				0;
		public Int16 cy =				0;
		public Int16 reservedMenu =		0;
		public Int16 reservedClass =	0;
		public Int16 reservedTitle =	0;

		// Single dlg item, must be dword-aligned - see documentation for DLGITEMTEMPLATE
		public Int32 itemStyle =			DlgStyle.WsChild;
		public Int32 itemExtendedStyle =	ExStyle.WsExNoParentNotify;
		public Int16 itemX =				0;
		public Int16 itemY =				0;
		public Int16 itemCx =				0;
		public Int16 itemCy =				0;
		public Int16 itemId =				0;
		public UInt16 itemClassHdr =		0xffff;	// we supply a constant to indicate the class of this control
		public Int16 itemClass =			0x0082;	// static label control
		public Int16 itemText =				0x0000;	// no text for this control
		public Int16 itemData =				0x0000;	// no creation data for this control
	};

	/// <summary>
	/// The rectangle structure used in Win32 API calls
	/// </summary>
	[StructLayout(LayoutKind.Sequential)]
	internal struct RECT 
	{
		public int left;
		public int top;
		public int right;
		public int bottom;
	};

	/// <summary>
	/// The point structure used in Win32 API calls
	/// </summary>
	[StructLayout(LayoutKind.Sequential)]
	internal struct POINT
	{
		public int X;
		public int Y;
	};
	
	/// <summary>
	/// Contains all of the p/invoke declarations for the Win32 APIs used in this sample
	/// </summary>
	public class NativeMethods
	{

		[DllImport("User32.dll", CharSet = CharSet.Unicode)]
		internal static extern IntPtr GetDlgItem( IntPtr hWndDlg, Int16 Id );

		[DllImport("User32.dll", CharSet = CharSet.Unicode)]
		internal static extern IntPtr GetParent( IntPtr hWnd );

		[DllImport("User32.dll", CharSet = CharSet.Unicode)]
		internal static extern IntPtr SetParent( IntPtr hWndChild, IntPtr hWndNewParent );
		
		[DllImport("User32.dll", CharSet = CharSet.Unicode)]
		internal static extern UInt32 SendMessage( IntPtr hWnd, UInt32 msg, UInt32 wParam, StringBuilder buffer );

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		internal static extern int GetWindowRect( IntPtr hWnd, ref RECT rc );

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		internal static extern int GetClientRect( IntPtr hWnd, ref RECT rc );

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		internal static extern bool ScreenToClient( IntPtr hWnd, ref POINT pt );

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		internal static extern bool MoveWindow( IntPtr hWnd, int X, int Y, int Width, int Height, bool repaint );

		[DllImport("ComDlg32.dll", CharSet = CharSet.Unicode)]
		internal static extern bool GetOpenFileName( ref OpenFileName ofn );

		[DllImport("ComDlg32.dll", CharSet = CharSet.Unicode)]
		internal static extern Int32 CommDlgExtendedError();
	}
}
