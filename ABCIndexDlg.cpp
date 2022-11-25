#include "stdafx.h"

HINSTANCE hInst = 0;
const char appName[] = "ABCIndex";
static char appAboutName[] = "About ABCIndex";
const size_t blockSize = 15 * 1024;

const COLORREF black = RGB(0, 0, 0);
const COLORREF white = RGB(255, 255, 255);
const COLORREF red = RGB(255, 0, 0);
const COLORREF green = RGB(0, 255, 0);
const COLORREF blue = RGB(0, 0, 255);

int NewHandler(size_t size)
{
	UNREFERENCED_PARAMETER(size);
	throw WinException("Out of memory");
}

// Forward declarations of functions included in this code module:
INT_PTR CALLBACK DialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK About(HWND, UINT, WPARAM, LPARAM);

INT WINAPI WinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ PSTR szCmdLine, _In_ int iCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(szCmdLine);
	UNREFERENCED_PARAMETER(iCmdShow);

	hInst = hInstance;
	_set_new_handler(&NewHandler);

	HWND hDialog = 0;

	hDialog = CreateDialog(hInst, MAKEINTRESOURCE(IDD_ABCIndex_DIALOG), 0, static_cast<DLGPROC>(DialogProc));
	if (!hDialog)
	{
		char buf[100];
		wsprintf(buf, "Error x%x", GetLastError());
		MessageBox(0, buf, "CreateDialog", MB_ICONEXCLAMATION | MB_OK);
		return 1;
	}
	// Attach icon to main dialog
	const HANDLE hbicon = LoadImage(GetModuleHandle(0), MAKEINTRESOURCE(IDR_MAINFRAME),
		IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), 0);
	// load big icon
	if (hbicon)
		SendMessage(hDialog, WM_SETICON, ICON_BIG, LPARAM(hbicon));

	// load small icon
	const HANDLE hsicon = LoadImage(GetModuleHandle(0), MAKEINTRESOURCE(IDR_MAINFRAME),
		IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);
	if (hsicon)
		SendMessage(hDialog, WM_SETICON, ICON_SMALL, LPARAM(hsicon));

	MSG  msg;
	while (GetMessage(&msg, NULL, 0, 0) > 0)
	{
		// only handle non-dialog messages here
		if (!IsDialogMessage(hDialog, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	return (int)msg.wParam;
}

ABCIndex::ABCIndex(HWND hwnd)
	: m_seEditor(hwnd, 0, FALSE), // scite edit window
	m_bClose(hwnd, IDOK),
	m_bLoad(hwnd, IDC_LOAD),
	m_bScan(hwnd, IDC_SCAN),
	m_bSave(hwnd, IDC_SAVE),
	m_bClear(hwnd, IDC_CLEAR)
{
	// Attach icon to main dialog
	HICON hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDR_MAINFRAME));
	SendMessage(hwnd, WM_SETICON, WPARAM(TRUE), LPARAM(hIcon));

	// Other initializations...
	hwndDialog = hwnd;
	fullPath[0] = '\0';
	isDirty = false;
	m_bSave.Enable(false);
	m_bScan.Enable(false);
	m_bClear.Enable(false);
#ifdef _USE_TUPLE
	// construct tuple vector
	// grave
	tTupleMap.push_back(make_tuple(u8"À", string("`A")));
	tTupleMap.push_back(make_tuple(u8"À", string("`A")));
	tTupleMap.push_back(make_tuple(u8"à", string("`a")));
	tTupleMap.push_back(make_tuple(u8"È", string("`E")));
	tTupleMap.push_back(make_tuple(u8"è", string("`e")));
	tTupleMap.push_back(make_tuple(u8"Ì", string("`I")));
	tTupleMap.push_back(make_tuple(u8"ì", string("`i")));
	tTupleMap.push_back(make_tuple(u8"Ò", string("`O")));
	tTupleMap.push_back(make_tuple(u8"ò", string("`o")));
	tTupleMap.push_back(make_tuple(u8"Ù", string("`U")));
	tTupleMap.push_back(make_tuple(u8"ù", string("`u")));
	// acute
	tTupleMap.push_back(make_tuple(u8"Á", string("'A")));
	tTupleMap.push_back(make_tuple(u8"á", string("'a")));
	tTupleMap.push_back(make_tuple(u8"É", string("'E")));
	tTupleMap.push_back(make_tuple(u8"é", string("'e")));
	tTupleMap.push_back(make_tuple(u8"Í", string("'I")));
	tTupleMap.push_back(make_tuple(u8"í", string("'i")));
	tTupleMap.push_back(make_tuple(u8"Ó", string("'O")));
	tTupleMap.push_back(make_tuple(u8"ó", string("'o")));
	tTupleMap.push_back(make_tuple(u8"Ú", string("'U")));
	tTupleMap.push_back(make_tuple(u8"ú", string("'u")));
	tTupleMap.push_back(make_tuple(u8"Ý", string("'Y")));
	tTupleMap.push_back(make_tuple(u8"ý", string("'y")));
	// circumflex
	tTupleMap.push_back(make_tuple(u8"Â", string("^A")));
	tTupleMap.push_back(make_tuple(u8"â", string("^a")));
	tTupleMap.push_back(make_tuple(u8"Ê", string("^E")));
	tTupleMap.push_back(make_tuple(u8"ê", string("^e")));
	tTupleMap.push_back(make_tuple(u8"Î", string("^I")));
	tTupleMap.push_back(make_tuple(u8"î", string("^i")));
	tTupleMap.push_back(make_tuple(u8"Ô", string("^O")));
	tTupleMap.push_back(make_tuple(u8"ô", string("^o")));
	tTupleMap.push_back(make_tuple(u8"Û", string("^U")));
	tTupleMap.push_back(make_tuple(u8"û", string("^u")));
	tTupleMap.push_back(make_tuple(u8"Ŷ", string("^Y")));
	tTupleMap.push_back(make_tuple(u8"ŷ", string("^y")));
	// tilde
	tTupleMap.push_back(make_tuple(u8"Ã", string("~A")));
	tTupleMap.push_back(make_tuple(u8"ã", string("~a")));
	tTupleMap.push_back(make_tuple(u8"Ñ", string("~N")));
	tTupleMap.push_back(make_tuple(u8"ñ", string("~n")));
	tTupleMap.push_back(make_tuple(u8"Õ", string("~O")));
	tTupleMap.push_back(make_tuple(u8"õ", string("~o")));
	// umlaut
	tTupleMap.push_back(make_tuple(u8"Ä", string("\"A")));
	tTupleMap.push_back(make_tuple(u8"ä", string("\"a")));
	tTupleMap.push_back(make_tuple(u8"Ë", string("\"E")));
	tTupleMap.push_back(make_tuple(u8"ë", string("\"e")));
	tTupleMap.push_back(make_tuple(u8"Ï", string("\"I")));
	tTupleMap.push_back(make_tuple(u8"ï", string("\"i")));
	tTupleMap.push_back(make_tuple(u8"Ö", string("\"O")));
	tTupleMap.push_back(make_tuple(u8"ö", string("\"o")));
	tTupleMap.push_back(make_tuple(u8"Ü", string("\"U")));
	tTupleMap.push_back(make_tuple(u8"ü", string("\"u")));
	tTupleMap.push_back(make_tuple(u8"Ÿ", string("\"Y")));
	tTupleMap.push_back(make_tuple(u8"ÿ", string("\"y")));
	// cedilla
	tTupleMap.push_back(make_tuple(u8"Ç", string("cC")));
	tTupleMap.push_back(make_tuple(u8"ç", string("cc")));
	tTupleMap.push_back(make_tuple(u8"Ç", string(",C")));
	tTupleMap.push_back(make_tuple(u8"ç", string(",c")));
	// ring
	tTupleMap.push_back(make_tuple(u8"Å", string("AA")));
	tTupleMap.push_back(make_tuple(u8"å", string("aa")));
	tTupleMap.push_back(make_tuple(u8"Å", string("oA")));
	tTupleMap.push_back(make_tuple(u8"å", string("oa")));
	// slash
	tTupleMap.push_back(make_tuple(u8"Ø", string("/O")));
	tTupleMap.push_back(make_tuple(u8"ø", string("/o")));
	// breve
	tTupleMap.push_back(make_tuple(u8"Ă", string("uA")));
	tTupleMap.push_back(make_tuple(u8"ă", string("ua")));
	tTupleMap.push_back(make_tuple(u8"Ĕ", string("uE")));
	tTupleMap.push_back(make_tuple(u8"ĕ", string("ue")));
	// caron
	tTupleMap.push_back(make_tuple(u8"Š", string("vS")));
	tTupleMap.push_back(make_tuple(u8"š", string("vs")));
	tTupleMap.push_back(make_tuple(u8"Ž", string("vZ")));
	tTupleMap.push_back(make_tuple(u8"ž", string("vz")));
	// double acute
	tTupleMap.push_back(make_tuple(u8"Ő", string("HO")));
	tTupleMap.push_back(make_tuple(u8"ő", string("Ho")));
	tTupleMap.push_back(make_tuple(u8"Ű", string("HU")));
	tTupleMap.push_back(make_tuple(u8"ű", string("Hu")));
	// ligatures
	tTupleMap.push_back(make_tuple(u8"Æ", string("AE")));
	tTupleMap.push_back(make_tuple(u8"æ", string("ae")));
	tTupleMap.push_back(make_tuple(u8"Œ", string("OE")));
	tTupleMap.push_back(make_tuple(u8"œ", string("oe")));
	tTupleMap.push_back(make_tuple(u8"ß", string("ss")));
	tTupleMap.push_back(make_tuple(u8"Ð", string("DH")));
	tTupleMap.push_back(make_tuple(u8"ð", string("dh")));
	tTupleMap.push_back(make_tuple(u8"Þ", string("TH")));
	tTupleMap.push_back(make_tuple(u8"þ", string("th")));
#else
	// construct map of pairs
	// grave
	tAccentMap.insert(make_pair(u8"À", string("`A")));
	tAccentMap.insert(make_pair(u8"À", string("`A")));
	tAccentMap.insert(make_pair(u8"à", string("`a")));
	tAccentMap.insert(make_pair(u8"È", string("`E")));
	tAccentMap.insert(make_pair(u8"è", string("`e")));
	tAccentMap.insert(make_pair(u8"Ì", string("`I")));
	tAccentMap.insert(make_pair(u8"ì", string("`i")));
	tAccentMap.insert(make_pair(u8"Ò", string("`O")));
	tAccentMap.insert(make_pair(u8"ò", string("`o")));
	tAccentMap.insert(make_pair(u8"Ù", string("`U")));
	tAccentMap.insert(make_pair(u8"ù", string("`u")));
	// acute
	tAccentMap.insert(make_pair(u8"Á", string("'A")));
	tAccentMap.insert(make_pair(u8"á", string("'a")));
	tAccentMap.insert(make_pair(u8"É", string("'E")));
	tAccentMap.insert(make_pair(u8"é", string("'e")));
	tAccentMap.insert(make_pair(u8"Í", string("'I")));
	tAccentMap.insert(make_pair(u8"í", string("'i")));
	tAccentMap.insert(make_pair(u8"Ó", string("'O")));
	tAccentMap.insert(make_pair(u8"ó", string("'o")));
	tAccentMap.insert(make_pair(u8"Ú", string("'U")));
	tAccentMap.insert(make_pair(u8"ú", string("'u")));
	tAccentMap.insert(make_pair(u8"Ý", string("'Y")));
	tAccentMap.insert(make_pair(u8"ý", string("'y")));
	// circumflex
	tAccentMap.insert(make_pair(u8"Â", string("^A")));
	tAccentMap.insert(make_pair(u8"â", string("^a")));
	tAccentMap.insert(make_pair(u8"Ê", string("^E")));
	tAccentMap.insert(make_pair(u8"ê", string("^e")));
	tAccentMap.insert(make_pair(u8"Î", string("^I")));
	tAccentMap.insert(make_pair(u8"î", string("^i")));
	tAccentMap.insert(make_pair(u8"Ô", string("^O")));
	tAccentMap.insert(make_pair(u8"ô", string("^o")));
	tAccentMap.insert(make_pair(u8"Û", string("^U")));
	tAccentMap.insert(make_pair(u8"û", string("^u")));
	tAccentMap.insert(make_pair(u8"Ŷ", string("^Y")));
	tAccentMap.insert(make_pair(u8"ŷ", string("^y")));
	// tilde
	tAccentMap.insert(make_pair(u8"Ã", string("~A")));
	tAccentMap.insert(make_pair(u8"ã", string("~a")));
	tAccentMap.insert(make_pair(u8"Ñ", string("~N")));
	tAccentMap.insert(make_pair(u8"ñ", string("~n")));
	tAccentMap.insert(make_pair(u8"Õ", string("~O")));
	tAccentMap.insert(make_pair(u8"õ", string("~o")));
	// umlaut
	tAccentMap.insert(make_pair(u8"Ä", string("\"A")));
	tAccentMap.insert(make_pair(u8"ä", string("\"a")));
	tAccentMap.insert(make_pair(u8"Ë", string("\"E")));
	tAccentMap.insert(make_pair(u8"ë", string("\"e")));
	tAccentMap.insert(make_pair(u8"Ï", string("\"I")));
	tAccentMap.insert(make_pair(u8"ï", string("\"i")));
	tAccentMap.insert(make_pair(u8"Ö", string("\"O")));
	tAccentMap.insert(make_pair(u8"ö", string("\"o")));
	tAccentMap.insert(make_pair(u8"Ü", string("\"U")));
	tAccentMap.insert(make_pair(u8"ü", string("\"u")));
	tAccentMap.insert(make_pair(u8"Ÿ", string("\"Y")));
	tAccentMap.insert(make_pair(u8"ÿ", string("\"y")));
	// cedilla
	tAccentMap.insert(make_pair(u8"Ç", string("cC")));
	tAccentMap.insert(make_pair(u8"ç", string("cc")));
	tAccentMap.insert(make_pair(u8"Ç", string(",C")));
	tAccentMap.insert(make_pair(u8"ç", string(",c")));
	// ring
	tAccentMap.insert(make_pair(u8"Å", string("AA")));
	tAccentMap.insert(make_pair(u8"å", string("aa")));
	tAccentMap.insert(make_pair(u8"Å", string("oA")));
	tAccentMap.insert(make_pair(u8"å", string("oa")));
	// slash
	tAccentMap.insert(make_pair(u8"Ø", string("/O")));
	tAccentMap.insert(make_pair(u8"ø", string("/o")));
	// breve
	tAccentMap.insert(make_pair(u8"Ă", string("uA")));
	tAccentMap.insert(make_pair(u8"ă", string("ua")));
	tAccentMap.insert(make_pair(u8"Ĕ", string("uE")));
	tAccentMap.insert(make_pair(u8"ĕ", string("ue")));
	// double acute
	tAccentMap.insert(make_pair(u8"Š", string("vS")));
	tAccentMap.insert(make_pair(u8"š", string("vs")));
	tAccentMap.insert(make_pair(u8"Ž", string("vZ")));
	tAccentMap.insert(make_pair(u8"ž", string("vz")));
	// caron
	tAccentMap.insert(make_pair(u8"Ő", string("HO")));
	tAccentMap.insert(make_pair(u8"ő", string("Ho")));
	tAccentMap.insert(make_pair(u8"Ű", string("HU")));
	tAccentMap.insert(make_pair(u8"ű", string("Hu")));
	// ligatures
	tAccentMap.insert(make_pair(u8"Æ", string("AE")));
	tAccentMap.insert(make_pair(u8"æ", string("ae")));
	tAccentMap.insert(make_pair(u8"Œ", string("OE")));
	tAccentMap.insert(make_pair(u8"œ", string("oe")));
	tAccentMap.insert(make_pair(u8"ß", string("ss")));
	tAccentMap.insert(make_pair(u8"Ð", string("DH")));
	tAccentMap.insert(make_pair(u8"ð", string("dh")));
	tAccentMap.insert(make_pair(u8"Þ", string("TH")));
	tAccentMap.insert(make_pair(u8"þ", string("th")));
#endif
	// set fixed font to make output columns neat
	// Set up the global default style. These attributes are used wherever no explicit choices are made.
	SetAStyle(STYLE_DEFAULT, black, white, 10, "courier new");
	m_seEditor.SendEditor(SCI_STYLECLEARALL);	// Copies global style to all others
}

void ABCIndex::SetAStyle(int style, COLORREF fore, COLORREF back, int size, const char* face)
{
	if (size >= 1)
		m_seEditor.SendEditor(SCI_STYLESETSIZE, style, size);
	if (face)
		m_seEditor.SendEditor(SCI_STYLESETFONT, style, reinterpret_cast<LPARAM>(face));
	m_seEditor.SendEditor(SCI_STYLESETFORE, style, fore);
	m_seEditor.SendEditor(SCI_STYLESETBACK, style, back);
}

void ABCIndex::Command(HWND hwnd, int controlID, int command)
{
	UNREFERENCED_PARAMETER(command);
	switch (controlID)
	{
	case IDC_LOAD:
		if (command == BN_CLICKED)
		{
			OnBnClickedBrowsefile();
		}
		break;
	case IDC_SCAN:
		if (command == BN_CLICKED)
		{
			OnBnClickedScanfile();
		}
		break;
	case IDC_SAVE:
		if (command == BN_CLICKED)
		{
			OnBnClickedSavefile();
		}
		break;
	case IDC_CLEAR:
		if (command == BN_CLICKED)
		{
			OnBnClickedClear();
		}
		break;
	case IDOK:
		PostMessage(hwnd, WM_CLOSE, 0, 0);
		break;
	}
}

void ABCIndex::OnBnClickedScanfile()
{
	if (FileList.size() > 0)
	{
		ProcessFiles();
		m_bScan.Enable(false);
		m_bSave.Enable(true);
	}
}

void ABCIndex::OnBnClickedSavefile()
{
	SaveAs();
}

void  ABCIndex::OnBnClickedClear()
{
	FileList.clear();
	ClearDisplay();
	m_bScan.Enable(false);
	m_bSave.Enable(false);
	m_bClear.Enable(false);
}

void ABCIndex::OnBnClickedBrowsefile()
{
	try
	{
		OPENFILENAME ofn;
		TCHAR szFile[1024]{ 0 };
		const int c_cMaxFiles = MAX_FILES;

		ZeroMemory(&ofn, sizeof(ofn));
		ofn.lStructSize = sizeof(ofn);
		ofn.lpstrFile = szFile;
		// Set lpstrFile[0] to '\0' so that GetOpenFileName does not
		// use the contents of szFile to initialize itself.
		ofn.lpstrFile[0] = '\0';
		ofn.hwndOwner = hwndDialog;
		ofn.nMaxFile = sizeof(szFile);
		ofn.lpstrFilter = TEXT("ABC\0*.abc\0";);  // Filter for ABC files
		ofn.nFilterIndex = 1;
		ofn.lpstrInitialDir = NULL;
		ofn.lpstrFileTitle = NULL;
		ofn.Flags = OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT;

		if (!GetOpenFileName(&ofn))
		{
			// GetOpenFileName returns 0 if there is an error or cancel pressed
			// CommDlgExtendedError returns 12291 if buffer is not large enough for all filenames selected
			DWORD errorword = CommDlgExtendedError();
			if (errorword)
			{
				throw (ErrorClass(SYSTEM_ERROR, "File selection error.", errorword, "CommDlgExtendedError:\n"));
			}
			return;
		}

		char* NextFileName = szFile;
		string File1, nextPath;

		// if multi-select was used, Path is a 0 separated list of dir and filenames.
		if (strlen(ofn.lpstrFile) + 1 != ofn.nFileOffset)
		{
			// Just one file was selected
			// File1 = string((ofn.lpstrFile + ofn.nFileOffset));
			nextPath = string(ofn.lpstrFile);
			// check against existing file count
			if (FileList.size() < c_cMaxFiles)
			{
				size_t result = AddUniqueFile(&nextPath);
				if (result != 255 && result)
				{
					u8string translated;
					ConvertLine(&FileList[result - 1], &translated);
					UpdateDisplay(&translated);
				}
			}
			else
			{
				throw (ErrorClass(SYSTEM_ERROR, "File selection error.", FileList.size(), "Too many files selected:\n"));
			}
		}
		else
		{
			// Multiple files were selected so extract them.
			// Again "strlen"'s behaviour to assume string is 0 terminated will be used here.

			// iterate through file names and count them
			NextFileName = ofn.lpstrFile + ofn.nFileOffset; //Do first iteration already to jump to first filename and skip dir
			int selFileCount = 0;
			while (*NextFileName)
			{
				File1 = string(NextFileName);
				nextPath = string(ofn.lpstrFile) + string("\\") + string(File1);;
				selFileCount++;
				NextFileName = NextFileName + File1.size() + 1;
			}
			// check that max files will not be exceeded
			DWORD totalCount = selFileCount + (DWORD)(FileList.size());
			if (totalCount > c_cMaxFiles)
			{
				throw (ErrorClass(SYSTEM_ERROR, "File selection error.", totalCount, "Too many files selected:\n"));
			}
			else
			{
				NextFileName = ofn.lpstrFile + ofn.nFileOffset; //Do first iteration already to jump to first filename and skip dir
				while (*NextFileName)
				{
					File1 = string(NextFileName); // strcpy_s(tempFile, sizeof(tempFile), szFile);
					nextPath = string(ofn.lpstrFile) + string("\\") + string(File1);;
					if (FileList.size() < c_cMaxFiles)
					{
						size_t result = AddUniqueFile(&nextPath);
						if (result != 255 && result)
						{
							UpdateDisplay(&FileList[result - 1]);
						}
					}
					else
					{
						throw (ErrorClass(SYSTEM_ERROR, "File selection error.", FileList.size(), "Too many files selected:\n"));
					}
					NextFileName = NextFileName + File1.size() + 1;
				}
			}
		}
		m_bScan.Enable(true);
		m_bClear.Enable(true);
	}
	catch (WinException e)
	{
		MessageBox(0, e.GetMessage(), "Exception", MB_ICONEXCLAMATION | MB_OK);
	}
	catch (ErrorClass error)
	{
		LPVOID lpMsgBuf = nullptr;
		LPVOID lpDisplayBuf;
		LPTSTR lpszFunction = error.ErrorString;
		LPCVOID lpSource;    // pointer to  message source
		DWORD dwMessageId;   // requested message identifier

		switch (error.ErrorType)
		{
		case SYSTEM_ERROR:          // error before file is opened
			lpSource = NULL;
			dwMessageId = error.ErrorWord;
			break;
		case APPLICATION_ERROR:
			lpSource = error.ErrorString;
			dwMessageId = error.ErrorWord;
			break;
		default:
			lpSource = "Unknown Error";
			dwMessageId = error.ErrorWord;
			break;
		}
		FormatMessage(
			FORMAT_MESSAGE_ALLOCATE_BUFFER |
			FORMAT_MESSAGE_FROM_SYSTEM |
			FORMAT_MESSAGE_IGNORE_INSERTS,	// source and processing options
			lpSource,						// pointer to  message source
			dwMessageId,					// requested message identifier
			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
			(LPTSTR)&lpMsgBuf,
			0,
			NULL);

		// Display the error message and exit the process
		lpDisplayBuf = (LPVOID)LocalAlloc(LMEM_ZEROINIT, ((size_t)lstrlen((LPCTSTR)lpMsgBuf) + (size_t)lstrlen((LPCTSTR)lpszFunction) + 40) * sizeof(TCHAR));
		if (lpDisplayBuf)
		{
			HRESULT hr = StringCchPrintf((LPTSTR)lpDisplayBuf, LocalSize(lpDisplayBuf) / sizeof(TCHAR), TEXT("%hs failed with error %d"), lpszFunction, dwMessageId);
			if (hr == S_OK)
			{
				MessageBox(hwndDialog, (LPCTSTR)lpDisplayBuf, TEXT(error.ErrorTitle), MB_ICONEXCLAMATION | MB_OK);
				LocalFree(lpMsgBuf);
				LocalFree(lpDisplayBuf);
			}
		}
	}
	catch (...)
	{
		MessageBox(0, "Unknown", "Exception", MB_ICONEXCLAMATION | MB_OK);
	}
}

// handle drag and drop files
void ABCIndex::OnDropFiles(HDROP hDropInfo)
{
	try
	{
		const int c_cMaxFiles = MAX_FILES;
		const int c_cbBuffSize = (c_cMaxFiles * (MAX_PATH + 1)) + 1;
		TCHAR filenameBuffer[c_cbBuffSize]{ 0 };
		// UINT_MAX means return the number of files dropped.
		int dropFileCount = DragQueryFile(hDropInfo, UINT_MAX, filenameBuffer, c_cbBuffSize);
		size_t totalCount = dropFileCount + FileList.size();
		// validate number of files selected
		if (totalCount > c_cMaxFiles)
		{
			throw (ErrorClass(SYSTEM_ERROR, "File selection error.", totalCount, "Too many files selected:\n"));
		}
		else
		{
			string nextPath;
			for (int index = 0; index < dropFileCount; index++)
			{
				// Grab the name of the file associated with index "i" in the list of files dropped.
				// Be sure you know that the name is attached to the FULL path of the file.
				DragQueryFile(hDropInfo, index, filenameBuffer, c_cbBuffSize);
				nextPath = string(filenameBuffer);
				size_t result = AddUniqueFile(&nextPath);
				if (result != 255 && result)
				{
					UpdateDisplay(&FileList[result - 1]);
				}
			}
			m_bScan.Enable(true);
		}
	}
	catch (ErrorClass error)
	{
		LPVOID lpMsgBuf = nullptr;
		LPVOID lpDisplayBuf;
		LPTSTR lpszFunction = error.ErrorString;
		LPCVOID lpSource;    // pointer to  message source
		DWORD dwMessageId;   // requested message identifier

		switch (error.ErrorType)
		{
		case SYSTEM_ERROR:   // error before file is opened
			lpSource = NULL;
			dwMessageId = error.ErrorWord;
			break;
		case APPLICATION_ERROR:
			lpSource = error.ErrorString;
			dwMessageId = error.ErrorWord;
			break;
		default:
			lpSource = "Unknown Error";
			dwMessageId = error.ErrorWord;
			break;
		}
		FormatMessage(
			FORMAT_MESSAGE_ALLOCATE_BUFFER |
			FORMAT_MESSAGE_FROM_SYSTEM |
			FORMAT_MESSAGE_IGNORE_INSERTS,        // source and processing options
			lpSource,        // pointer to  message source
			dwMessageId,     // requested message identifier
			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
			(LPTSTR)&lpMsgBuf,
			0,
			NULL);

		// Display the error message and exit the process
		lpDisplayBuf = (LPVOID)LocalAlloc(LMEM_ZEROINIT, ((size_t)lstrlen((LPCTSTR)lpMsgBuf) + (size_t)lstrlen((LPCTSTR)lpszFunction) + 40) * sizeof(TCHAR));
		if (lpDisplayBuf)
		{
			HRESULT hr = StringCchPrintf((LPTSTR)lpDisplayBuf, LocalSize(lpDisplayBuf) / sizeof(TCHAR), TEXT("%hs failed with error %d"), lpszFunction, dwMessageId);
			if (hr == S_OK)
			{
				MessageBox(hwndDialog, (LPCTSTR)lpDisplayBuf, TEXT(error.ErrorTitle), MB_ICONEXCLAMATION | MB_OK);
				LocalFree(lpMsgBuf);
				LocalFree(lpDisplayBuf);
			}
		}
	}
	catch (...)
	{
		MessageBox(0, "Unknown", "Exception", MB_ICONEXCLAMATION | MB_OK);
	}
}

// add a unit to the vector container but exclude duplicates
size_t ABCIndex::AddUniqueFile(string* fileName)
{
	if (FileList.size() >= MAX_FILES)
		return 255;
	string newFile = string(*fileName);
	if (find(FileList.begin(), FileList.end(), newFile) == FileList.end())
	{
		FileList.push_back(newFile);
		return FileList.size();
	}
	return 0;
}

bool ABCIndex::ProcessFiles()
{
	CMemMapFile mmf;  // instance of 'memorymapfile' by P J Naughter
	LPVOID pData;
	string filename;
	bool flag(false);
	HANDLE fileIn;
	BY_HANDLE_FILE_INFORMATION bhfi;
	ClearDisplay();
	for (UINT i = 0; i < FileList.size(); i++)
	{
		fileIn = INVALID_HANDLE_VALUE;
		filename = FileList[i];
		if ((fileIn = CreateFile(filename.c_str(),
			GENERIC_READ,
			0,
			0,
			OPEN_EXISTING,
			FILE_ATTRIBUTE_READONLY,
			0)) == INVALID_HANDLE_VALUE)
		{
			exit(1);
		}
		if (!GetFileInformationByHandle(fileIn, &bhfi))
		{
			CloseHandle(fileIn);
			exit(2);
		}
		CloseHandle(fileIn);

		mmf.MapFile(filename.c_str());
		pData = mmf.Open();

		//	make a string from the memory mapped file
		string ABCFile(static_cast<LPSTR>(pData), static_cast<DWORD>(bhfi.nFileSizeLow));
		string separator("---------------------------------------------------------------------------------------");
		// print filename header
		UpdateDisplay(&separator);
		u8string translated;
		ConvertLine(&filename, &translated);
		UpdateDisplay(&translated);
		// find tune number followed by tune name (X: ...T:
		regex index_regex("^X:[\\s\\S]*?^T:.*");
		// now find tune name field
		string textToken("T:");
		auto words_begin = sregex_iterator(ABCFile.begin(), ABCFile.end(), index_regex);
		auto words_end = sregex_iterator();

		for (sregex_iterator it = words_begin; it != words_end; ++it)
		{
			smatch match = *it;
			string match_str = match.str();
			string trimmed;
			size_t found = match_str.find(textToken);
			if (found != string::npos)
			{
				string output = regex_replace(match_str.substr(0, found), regex("[^0-9]*"), string("$1")); // get tune number
				match_str.erase(0, found + 2);  // remove up to start of T: field
				trimmed = ltrim(match_str);		// trim leading white space
				output = output + string(".\t") + trimmed;
				u8string accented;
				ConvertLine(&output, &accented);
				UpdateDisplay(&accented);
			}
		}
		mmf.UnMap();
	}
	if (FileList.size())
		flag = true;
	return flag;
}

void ABCIndex::Save()
{
	SaveFile(fullPath);
}

void ABCIndex::SaveAs()
{
	char openName[MAX_PATH] = "\0";
	strcpy_s(openName, fullPath);
	OPENFILENAME ofn = { sizeof(ofn) };
	ofn.hwndOwner = hwndDialog;
	ofn.hInstance = hInst;
	ofn.lpstrFile = openName;
	ofn.nMaxFile = sizeof(openName);
	ofn.lpstrTitle = "Save File";
	ofn.Flags = OFN_HIDEREADONLY;

	if (::GetSaveFileName(&ofn))
	{
		strcpy_s(fullPath, openName);
		SaveFile(fullPath);
	}
}

void ABCIndex::SaveFile(const char* fileName)
{
	UNREFERENCED_PARAMETER(fileName);
	FILE* fp;
	errno_t err = fopen_s(&fp, fullPath, "wb");
	if (err == 0 && fp)
	{
		char data[blockSize + 1];
		size_t lengthDoc = (size_t)m_seEditor.SendEditor(SCI_GETLENGTH);
		for (size_t index = 0; index < lengthDoc; index += blockSize)
		{
			size_t grabSize = lengthDoc - index;
			if (grabSize > blockSize)
				grabSize = blockSize;
			GetRange(index, index + grabSize, data);
			fwrite(data, grabSize, 1, fp);
		}
		fclose(fp);
		m_seEditor.SendEditor(SCI_SETSAVEPOINT);
	}
	else
	{
		char msg[MAX_PATH + 100];
		strcpy_s(msg, "Could not save file \"");
		strcat_s(msg, fullPath);
		strcat_s(msg, "\".");
		MessageBox(hwndDialog, msg, appName, MB_OK);
	}
}

int ABCIndex::SaveIfUnsure()
{
	if (isDirty)
	{
		char msg[MAX_PATH + 100];
		strcpy_s(msg, "Save changes to \"");
		strcat_s(msg, fullPath);
		strcat_s(msg, "\"?");
		int decision = MessageBox(hwndDialog, msg, appName, MB_YESNOCANCEL);
		if (decision == IDYES)
		{
			Save();
		}
		return decision;
	}
	return IDYES;
}

void ABCIndex::GetRange(size_t start, size_t end, char* text)
{
	TEXTRANGE tr{ 0 };
	tr.chrg.cpMin = (LONG)start;
	tr.chrg.cpMax = (LONG)end;
	tr.lpstrText = text;
	m_seEditor.SendEditor(EM_GETTEXTRANGE, 0, reinterpret_cast<LPARAM>(&tr));
}

std::string ltrim(const std::string& s)
{
	const std::string WHITESPACE = " \n\r\t\f\v";
	size_t start = s.find_first_not_of(WHITESPACE);
	return (start == std::string::npos) ? "" : s.substr(start);
}

// clear the scintilla control
void ABCIndex::ClearDisplay()
{
	m_seEditor.SendEditor(SCI_CLEARALL);
	m_seEditor.Enable(FALSE);
}

// write text to the scintilla control
void ABCIndex::UpdateDisplay(std::u8string* ostr)
{
	m_seEditor.SendEditor(SCI_APPENDTEXT, ostr->length(), (LPARAM)ostr->c_str());
	m_seEditor.SendEditor(SCI_APPENDTEXT, 1, (LPARAM)L"\n\r");
	m_seEditor.SendEditor(SCI_SCROLLCARET);
	m_seEditor.Enable(TRUE);
}

void ABCIndex::UpdateDisplay(std::string* ostr)
{
	m_seEditor.SendEditor(SCI_APPENDTEXT, ostr->length(), (LPARAM)ostr->c_str());
	m_seEditor.SendEditor(SCI_APPENDTEXT, 1, (LPARAM)L"\n\r");
	m_seEditor.SendEditor(SCI_SCROLLCARET);
	m_seEditor.Enable(TRUE);
}

// look at string for escape sequences
bool ABCIndex::TranslateEscape(size_t i, std::string* ostr, std::u8string* accented)
{
	bool found = false;
	if (i < ostr->length() - 2)
	{
		// get the two characters after the backslash into a string
		string modifier = string(&ostr->at(i + 1), 2);
		u8string acute;
#ifdef _USE_TUPLE
		// look for a match in the tuples vector
		for (auto const& it : tTupleMap)
		{
			if (get<1>(it) == modifier)
			{
				acute = get<0>(it);
				accented->assign(acute);
				found = true;
			}
		}
#else
		// look for a match in the characters map
		for (auto it = tAccentMap.begin(); it != tAccentMap.end(); ++it)
		{
			if (it->second == modifier)
			{
				acute = it->first;
				accented->assign(acute);
				found = true;
			}
		}
#endif
	}
	return found;
}

// source is ostr translated string is in translated
void ABCIndex::ConvertLine(std::string* ostr, std::u8string* translated)
{
	unsigned char character;
	size_t i = 0;
	// examine each character looking for escaped sequence like "\'a"
	while (i < ostr->length())
	{
		u8string acute;
		character = ostr->at(i);
		if (character > 0x7f) // first look for upper single byte characters and convert them to utf-8
		{
			if (((character & 0xe0) == 0xc0) && (ostr->at(i + 1) & 0xc0) == 0x80)
			{
				// already utf-8
				translated->push_back(character);
				translated->push_back(ostr->at(i + 1));
				i += 2; // step over utf-8 2 byte character
			}
			else
			{
				translated->push_back(0xc0 | character >> 6);
				translated->push_back(0x80 | (character & 0x3f));
				i++;
			}
		}
		else if (character == ('\\')) // // now look for escape sequence
		{
			if (TranslateEscape(i, ostr, &acute))
			{
				// escaped character found so acute contains translated character
				i += 2; // step over escaped characters
			}
			else
			{
				// not escape character so print backslash and continue
				acute = character;
			}
			translated->append(acute);
			i++;
		}
		else if (character == ('{') || character == ('}')) // now look for escape sequence
		{
			// ignore braces (old way of defining special characters
			i++;
		}
		else // not a special character
		{
			acute = character;
			translated->append(acute);
			i++;
		}
	}
}

INT_PTR CALLBACK DialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);

	static ABCIndex* control = 0;
	switch (message)
	{
	case WM_INITDIALOG:
		try
		{
			control = new ABCIndex(hwnd);
			HMENU hSysMenu;
			hSysMenu = GetSystemMenu(hwnd, FALSE);
			// add a system menu separator
			AppendMenu(hSysMenu, MF_SEPARATOR, NULL, NULL);
			// add a system menu item
			AppendMenu(hSysMenu, MF_STRING, IDS_ABOUTBOX, appAboutName);
		}
		catch (WinException e)
		{
			MessageBox(0, e.GetMessage(), "Exception", MB_ICONEXCLAMATION | MB_OK);
		}
		catch (...)
		{
			MessageBox(0, "Unknown", "Exception", MB_ICONEXCLAMATION | MB_OK);
			return -1;
		}
		return TRUE;
	case WM_COMMAND:
		if (control)
			control->Command(hwnd, LOWORD(wParam), HIWORD(wParam));
		return TRUE;
	case WM_DESTROY:
		PostQuitMessage(0);
		return TRUE;
	case WM_CLOSE:
		if (control)
			delete control;
		DestroyWindow(hwnd);
		return TRUE;
	case WM_SYSCOMMAND:
	{
		switch (wParam)
		{
		case IDS_ABOUTBOX:
		{
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hwnd, About);
			break;
		}
		}
		break;
	}
	case WM_DROPFILES:
	{
		HDROP fDrop = (HDROP)wParam;
		if (control)
			control->OnDropFiles(fDrop);
	}

	}
	return FALSE;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);

	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}

