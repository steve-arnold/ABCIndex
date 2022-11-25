#pragma once

#include <windows.h>
#include <map>
#include <regex>
#include <string>
#include <strsafe.h>
#include <richedit.h>

#include <vector>
#include <tuple>

#include "controls.h"
#include "ABCIndex_add.h"
using namespace std;
#define _USE_TUPLE

constexpr auto MAX_FILES = 20;

class WinException
{
public:
	WinException(const char* msg)
		: _err(GetLastError()), _msg(msg) {}
	DWORD GetError() const { return _err; }
	const char* GetMessage() const { return _msg; }
private:
	DWORD _err;
	const char* _msg;
};

// out of memory handler that throws WinException
int NewHandler(size_t size);

class ABCIndex
{
public:
	ABCIndex(HWND hwnd);
	~ABCIndex() {}
	void Command(HWND hwnd, int controlID, int command);
	void OnBnClickedBrowsefile();
	void OnBnClickedScanfile();
	void OnBnClickedSavefile();
	void OnBnClickedClear();
	void OnDropFiles(HDROP hDropInfo);
	size_t AddUniqueFile(string* fileName);
	void ClearDisplay();
	void UpdateDisplay(std::u8string* ostr);
	void UpdateDisplay(std::string* ostr);
	void ConvertLine(std::string* ostr, std::u8string* accented);
	bool ProcessFiles();
	void Save();
	void SaveAs();
	void SaveFile(const char* fileName);
	int SaveIfUnsure();
	void GetRange(size_t start, size_t end, char* text);
	void SetAStyle(int style, COLORREF fore, COLORREF back, int size, const char* face);
	bool TranslateEscape(size_t i, std::string* ostr, std::u8string* accented);
	bool isDirty;
	char fullPath[MAX_PATH];
private:
	HWND        hwndDialog;
	SciEdit     m_seEditor;
	Button      m_bLoad;
	Button      m_bClose;
	Button		m_bScan;
	Button      m_bSave;
	Button      m_bClear;
#ifdef _USE_TUPLE
	vector<tuple<u8string, std::string>> tTupleMap;
#else
	map<u8string, std::string> tAccentMap;
#endif
	vector<string> FileList; // container for list of program units

};
std::string ltrim(const std::string& s);

class ErrorClass
{
public:
	ErrorClass(int err, const char* title, size_t errorword, const char* string)
	{
		ErrorType = err;
		ErrorWord = (DWORD)(errorword & 0xffff);
		strcpy_s(ErrorTitle, title);
		strcpy_s(ErrorString, string);
	};

	~ErrorClass()
	{

	};
	int ErrorType;
	DWORD ErrorWord;
	char ErrorString[255];
	char ErrorTitle[255];
};
enum errorclass { SYSTEM_ERROR = 1, SYSTEM_ERROR2, APPLICATION_ERROR, APPLICATION_ERROR2 };
