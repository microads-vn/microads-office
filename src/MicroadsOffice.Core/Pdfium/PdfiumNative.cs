using System.Runtime.InteropServices;

namespace MicroadsOffice.Core.Pdfium;

/// <summary>P/Invoke bindings for the PDFium engine (pdfium.dll shipped next to the executable).</summary>
internal static class PdfiumNative
{
    private const string Dll = "pdfium.dll";

    public const int FPDF_ANNOT = 0x01;
    public const int FPDF_LCD_TEXT = 0x02;
    public const int FPDF_NO_NATIVETEXT = 0x04;
    public const int FPDF_GRAYSCALE = 0x08;
    public const int FPDF_PRINTING = 0x800;
    public const int FPDF_REVERSE_BYTE_ORDER = 0x10;

    public const int FPDFBitmap_BGRA = 4;
    public const int FPDFBitmap_BGRx = 3;

    public const int FPDF_ERR_SUCCESS = 0, FPDF_ERR_UNKNOWN = 1, FPDF_ERR_FILE = 2, FPDF_ERR_FORMAT = 3,
        FPDF_ERR_PASSWORD = 4, FPDF_ERR_SECURITY = 5, FPDF_ERR_PAGE = 6;

    public const uint PDFACTION_UNSUPPORTED = 0, PDFACTION_GOTO = 1, PDFACTION_REMOTEGOTO = 2, PDFACTION_URI = 3, PDFACTION_LAUNCH = 4;

    public const int FPDF_MATCHCASE = 0x1, FPDF_MATCHWHOLEWORD = 0x2;

    [StructLayout(LayoutKind.Sequential)]
    public struct FS_SIZEF { public float Width; public float Height; }

    [StructLayout(LayoutKind.Sequential)]
    public struct FPDF_LIBRARY_CONFIG
    {
        public int version;
        public IntPtr m_pUserFontPaths;
        public IntPtr m_pIsolate;
        public uint m_v8EmbedderSlot;
        public IntPtr m_pPlatform;
        public int m_RendererType;
    }

    [DllImport(Dll)] public static extern void FPDF_InitLibraryWithConfig(ref FPDF_LIBRARY_CONFIG config);
    [DllImport(Dll)] public static extern void FPDF_InitLibrary();
    [DllImport(Dll)] public static extern IntPtr FPDF_LoadMemDocument64(IntPtr dataBuf, nuint size, [MarshalAs(UnmanagedType.LPUTF8Str)] string? password);
    [DllImport(Dll)] public static extern void FPDF_CloseDocument(IntPtr document);
    [DllImport(Dll)] public static extern uint FPDF_GetLastError();
    [DllImport(Dll)] public static extern int FPDF_GetPageCount(IntPtr document);
    [DllImport(Dll)] public static extern int FPDF_GetPageSizeByIndexF(IntPtr document, int pageIndex, out FS_SIZEF size);
    [DllImport(Dll)] public static extern IntPtr FPDF_LoadPage(IntPtr document, int pageIndex);
    [DllImport(Dll)] public static extern void FPDF_ClosePage(IntPtr page);
    [DllImport(Dll)] public static extern int FPDF_GetFileVersion(IntPtr doc, out int fileVersion);
    [DllImport(Dll)] public static extern uint FPDF_GetDocPermissions(IntPtr document);
    [DllImport(Dll)] public static extern int FPDF_GetSecurityHandlerRevision(IntPtr document);
    [DllImport(Dll)] public static extern int FPDFPage_GetRotation(IntPtr page);
    [DllImport(Dll)] public static extern float FPDF_GetPageWidthF(IntPtr page);
    [DllImport(Dll)] public static extern float FPDF_GetPageHeightF(IntPtr page);

    [DllImport(Dll)] public static extern IntPtr FPDFBitmap_CreateEx(int width, int height, int format, IntPtr firstScan, int stride);
    [DllImport(Dll)] public static extern void FPDFBitmap_Destroy(IntPtr bitmap);
    [DllImport(Dll)] public static extern int FPDFBitmap_FillRect(IntPtr bitmap, int left, int top, int width, int height, uint color);
    [DllImport(Dll)] public static extern void FPDF_RenderPageBitmap(IntPtr bitmap, IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, int flags);
    [DllImport(Dll)] public static extern int FPDF_DeviceToPage(IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, int deviceX, int deviceY, out double pageX, out double pageY);
    [DllImport(Dll)] public static extern int FPDF_PageToDevice(IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, double pageX, double pageY, out int deviceX, out int deviceY);

    // Text
    [DllImport(Dll)] public static extern IntPtr FPDFText_LoadPage(IntPtr page);
    [DllImport(Dll)] public static extern void FPDFText_ClosePage(IntPtr textPage);
    [DllImport(Dll)] public static extern int FPDFText_CountChars(IntPtr textPage);
    [DllImport(Dll)] public static extern uint FPDFText_GetUnicode(IntPtr textPage, int index);
    [DllImport(Dll)] public static extern int FPDFText_GetCharBox(IntPtr textPage, int index, out double left, out double right, out double bottom, out double top);
    [DllImport(Dll)] public static extern int FPDFText_GetCharIndexAtPos(IntPtr textPage, double x, double y, double xTolerance, double yTolerance);
    [DllImport(Dll)] public static extern int FPDFText_GetText(IntPtr textPage, int startIndex, int count, [Out] ushort[] result);
    [DllImport(Dll)] public static extern int FPDFText_CountRects(IntPtr textPage, int startIndex, int count);
    [DllImport(Dll)] public static extern int FPDFText_GetRect(IntPtr textPage, int rectIndex, out double left, out double top, out double right, out double bottom);
    [DllImport(Dll, CharSet = CharSet.Unicode)] public static extern IntPtr FPDFText_FindStart(IntPtr textPage, [MarshalAs(UnmanagedType.LPWStr)] string findWhat, uint flags, int startIndex);
    [DllImport(Dll)] public static extern int FPDFText_FindNext(IntPtr handle);
    [DllImport(Dll)] public static extern int FPDFText_GetSchResultIndex(IntPtr handle);
    [DllImport(Dll)] public static extern int FPDFText_GetSchCount(IntPtr handle);
    [DllImport(Dll)] public static extern void FPDFText_FindClose(IntPtr handle);

    // Links, bookmarks, destinations, actions
    [DllImport(Dll)] public static extern IntPtr FPDFLink_GetLinkAtPoint(IntPtr page, double x, double y);
    [DllImport(Dll)] public static extern IntPtr FPDFLink_GetDest(IntPtr document, IntPtr link);
    [DllImport(Dll)] public static extern IntPtr FPDFLink_GetAction(IntPtr link);
    [DllImport(Dll)] public static extern IntPtr FPDFBookmark_GetFirstChild(IntPtr document, IntPtr bookmark);
    [DllImport(Dll)] public static extern IntPtr FPDFBookmark_GetNextSibling(IntPtr document, IntPtr bookmark);
    [DllImport(Dll)] public static extern uint FPDFBookmark_GetTitle(IntPtr bookmark, byte[]? buffer, uint buflen);
    [DllImport(Dll)] public static extern IntPtr FPDFBookmark_GetDest(IntPtr document, IntPtr bookmark);
    [DllImport(Dll)] public static extern IntPtr FPDFBookmark_GetAction(IntPtr bookmark);
    [DllImport(Dll)] public static extern uint FPDFAction_GetType(IntPtr action);
    [DllImport(Dll)] public static extern IntPtr FPDFAction_GetDest(IntPtr document, IntPtr action);
    [DllImport(Dll)] public static extern uint FPDFAction_GetURIPath(IntPtr document, IntPtr action, byte[]? buffer, uint buflen);
    [DllImport(Dll)] public static extern int FPDFDest_GetDestPageIndex(IntPtr document, IntPtr dest);
    [DllImport(Dll)] public static extern int FPDFDest_GetLocationInPage(IntPtr dest, out int hasX, out int hasY, out int hasZoom, out float x, out float y, out float zoom);

    // Metadata
    [DllImport(Dll)] public static extern uint FPDF_GetMetaText(IntPtr document, [MarshalAs(UnmanagedType.LPStr)] string tag, byte[]? buffer, uint buflen);
    [DllImport(Dll)] public static extern uint FPDF_GetPageLabel(IntPtr document, int pageIndex, byte[]? buffer, uint buflen);

    // Forms (render interactive form fields)
    [DllImport(Dll)] public static extern IntPtr FPDFDOC_InitFormFillEnvironment(IntPtr document, IntPtr formInfo);
    [DllImport(Dll)] public static extern void FPDFDOC_ExitFormFillEnvironment(IntPtr formHandle);
    [DllImport(Dll)] public static extern void FORM_OnAfterLoadPage(IntPtr page, IntPtr formHandle);
    [DllImport(Dll)] public static extern void FORM_OnBeforeClosePage(IntPtr page, IntPtr formHandle);
    [DllImport(Dll)] public static extern void FPDF_FFLDraw(IntPtr formHandle, IntPtr bitmap, IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, int flags);
}
