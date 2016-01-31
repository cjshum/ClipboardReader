#define UNICODE
#define _UNICODE

#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "sapi.lib")
#pragma warning(disable : 4996)

#include <Windows.h>
#include <tchar.h>
#include <CommCtrl.h>
#include <sapi.h>
#include <sphelper.h>

//***************Variables and Definitions***************
//Tabs and Texboxes
enum TabIndexes
{
	ID_ReadingPageTab,
	ID_AutoClipboardTab
};
HACCEL SelectAll;
HWND TabControls;
HWND TextBox[2];
int ReadingPageTabIndex = 0;
int AutoClipboardTabIndex = 1;
TCHAR* TextBuffer = NULL;
unsigned int TextLength;

//Buttons
enum ReadStates
{
	StoppedState,
	ReadingState,
	PausedState
};
TCHAR* ReadPauseResumeText[3] =
{
	_T("Read"),
	_T("Pause"),
	_T("Resume"),
};
HWND ReadPauseResumeButton;
HWND StopButton;
int ReadState = StoppedState;

//Voice Configuration Controls
HWND VoiceList;
HWND VoiceRateSlider;
long VoiceRate;

//Checkboxes
HWND AutoClipboardCheckbox;
HWND AutoReadCheckbox;

//Others
#define WindowWidth 640
#define WindowHeight 480
enum CommandControls
{
	ID_SelectAll,
	ID_TabControls,
	ID_TextBox1,
	ID_TextBox2,
	ID_ReadPauseResumeButton,
	ID_StopButton,
	ID_VoiceList,
	ID_VoiceRate,
	ID_AutoClipboardCheckbox,
	ID_AutoReadCheckbox
};
HWND MainWindowHandle;
HWND ClipboardViewer;
ISpVoice* TextReader = NULL;
//***************Variables and Definitions***************

//Text Reader Functions
void ReadText()
{
	SendMessage(TextBox[ReadingPageTabIndex], EM_SETREADONLY, TRUE, 0);
	TextBuffer = new TCHAR[++TextLength];
	SendMessage(TextBox[ReadingPageTabIndex], WM_GETTEXT, TextLength, (LPARAM)TextBuffer);
	TextReader->Speak(TextBuffer, SPF_ASYNC|SPF_PURGEBEFORESPEAK, NULL);
	delete[] TextBuffer;
}

void TextReaderEvents()
{
	CSpEvent event;

	while (event.GetFrom(TextReader) == S_OK)
	{
		switch(event.eEventId)
		{
		case SPEI_END_INPUT_STREAM:
			if (ReadState != StoppedState)
			{
				if (SendMessage(AutoReadCheckbox, BM_GETSTATE, 0, 0) == BST_CHECKED)
				{
					TextLength = SendMessage(TextBox[AutoClipboardTabIndex], WM_GETTEXTLENGTH, 0, 0);
					if (TextLength > 0)
					{
						SendMessage(TextBox[ReadingPageTabIndex], EM_SETREADONLY, FALSE, 0);
						SendMessage(TextBox[ReadingPageTabIndex], WM_SETTEXT, 0, NULL);
					
						AutoClipboardTabIndex = ReadingPageTabIndex;
						ReadingPageTabIndex = !ReadingPageTabIndex;

						InvalidateRect(MainWindowHandle, NULL, TRUE);
						ReadText();
						break;
					}
				}
				SendMessage(TextBox[ReadingPageTabIndex], WM_SETTEXT, 0, NULL);
				SendMessage(MainWindowHandle, WM_COMMAND, ID_StopButton, 0);
			}
			break;

		case SPEI_WORD_BOUNDARY:
			{
				SPVOICESTATUS VoiceStatus;
				WPARAM nStart;
				LPARAM nEnd;
				TextReader->GetStatus(&VoiceStatus, NULL);
				nStart = (LPARAM)(VoiceStatus.ulInputWordPos / sizeof(char));
				nEnd = nStart + VoiceStatus.ulInputWordLen;
				if (IsWindowVisible(TextBox[ReadingPageTabIndex]))
					SetFocus(TextBox[ReadingPageTabIndex]);
				SendMessage(TextBox[ReadingPageTabIndex], EM_SETSEL, nStart, nEnd);
				SendMessage(TextBox[ReadingPageTabIndex], EM_SCROLLCARET, 0, 0);
			}
			break;
		}
	}
}

//Window Control Functions
WNDPROC OldTextBoxProcedure;
LRESULT CALLBACK NewTextBoxProcedure
	(
		HWND hwnd,
		UINT uMsg,
		WPARAM wParam,
		LPARAM lParam
	)
{
	WPARAM StartOfSelectection;
	LPARAM EndOfSelection;
	int SelectionSize;
	int TextSize;
	TCHAR* hMem1;
	TCHAR* hMem2;
	HGLOBAL ClipboardData;

	if (uMsg == WM_COPY)
	{
		OpenClipboard(NULL);

		SendMessage(hwnd, EM_GETSEL, (WPARAM)&StartOfSelectection, (LPARAM)&EndOfSelection);
		SelectionSize = EndOfSelection - StartOfSelectection + 1;

		TextSize = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0);
		TextSize -= (TextSize - EndOfSelection - 1);
		hMem1 = new TCHAR[TextSize]();
		hMem1[TextSize - 1] = NULL; //null terminated string
		SendMessage(hwnd, WM_GETTEXT, TextSize, (LPARAM)hMem1);

		ClipboardData = GlobalAlloc(GMEM_MOVEABLE, SelectionSize*sizeof(TCHAR));
		hMem2 = (TCHAR*)GlobalLock(ClipboardData);
		hMem2[SelectionSize - 1] = NULL; //null terminated string
		memcpy(hMem2, hMem1+StartOfSelectection, SelectionSize*sizeof(TCHAR));
		GlobalUnlock(ClipboardData);

		SetClipboardData(CF_UNICODETEXT, ClipboardData);

		CloseClipboard();
		delete[] hMem1;
		return 0;
	}
	return OldTextBoxProcedure(hwnd, uMsg, wParam, lParam);
}

void CreateControls(HWND hWndParent)
{
	TextReader->SetNotifyWindowMessage(hWndParent, WM_APP, 0, 0);

	const ACCEL AcceleratorStructure =
	{
		FCONTROL|FVIRTKEY, //fVirt
		0x41, //key
		ID_SelectAll //cmd
	};
	SelectAll = CreateAcceleratorTable((LPACCEL)&AcceleratorStructure, 1);

	const INITCOMMONCONTROLSEX CommonControls =
		{
			sizeof(INITCOMMONCONTROLSEX), //dwSize;
			ICC_TAB_CLASSES //dwICC;
		};
	InitCommonControlsEx(&CommonControls);
	TabControls = CreateWindow
	(
		WC_TABCONTROL, //lpClassName
		NULL, //lpWindowName
		WS_VISIBLE | WS_CHILD, //dwStyle
		0, //x
		0, //y
		WindowWidth, //nWidth
		WindowHeight -100, //nHeight
		hWndParent, //hWndParent
		(HMENU)ID_TabControls, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);
	TCITEM TabItemStructure = {0};
	TabItemStructure.mask = TCIF_TEXT;
	TabItemStructure.pszText = _T("Reading Page");
	SendMessage(TabControls, TCM_INSERTITEM, ID_ReadingPageTab, (LPARAM)&TabItemStructure);
	TabItemStructure.pszText = _T("Auto-Clipboard");
	SendMessage(TabControls, TCM_INSERTITEM, ID_AutoClipboardTab, (LPARAM)&TabItemStructure);

	TextBox[0] = CreateWindow
	(
		WC_EDIT, //lpClassName
		NULL, //lpWindowName
		WS_VISIBLE | WS_CHILD | WS_VSCROLL | WS_BORDER | ES_MULTILINE, //dwStyle
		5, //x
		30, //y
		WindowWidth -10, //nWidth
		WindowHeight -100-35, //nHeight
		TabControls, //hWndParent
		(HMENU)ID_TextBox1, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);
	SendMessage(TextBox[0], EM_SETLIMITTEXT, (WPARAM)-1, 0);
	OldTextBoxProcedure = (WNDPROC)SetWindowLongPtr(TextBox[0], GWLP_WNDPROC, (LONG)NewTextBoxProcedure);
	
	TextBox[1] = CreateWindow
	(
		WC_EDIT, //lpClassName
		NULL, //lpWindowName
		WS_CHILD | WS_VSCROLL | WS_BORDER | ES_MULTILINE, //dwStyle
		5, //x
		30, //y
		WindowWidth -10, //nWidth
		WindowHeight -100-35, //nHeight
		TabControls, //hWndParent
		(HMENU)ID_TextBox2, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);
	SendMessage(TextBox[1], EM_SETLIMITTEXT, (WPARAM)-1, 0);
	OldTextBoxProcedure = (WNDPROC)SetWindowLongPtr(TextBox[1], GWLP_WNDPROC, (LONG)NewTextBoxProcedure);

	ReadPauseResumeButton = CreateWindow
	(
		WC_BUTTON, //lpClassName
		_T("Read"), //lpWindowName
		BS_PUSHBUTTON | WS_VISIBLE | WS_CHILD, //dwStyle
		WindowWidth/2 -150/2 -150 -10, //x
		WindowHeight -100 -35, //y
		150, //nWidth
		25, //nHeight
		hWndParent, //hWndParent
		(HMENU)ID_ReadPauseResumeButton, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);

	StopButton = CreateWindow
	(
		WC_BUTTON, //lpClassName
		_T("Stop"), //lpWindowName
		BS_PUSHBUTTON | WS_VISIBLE | WS_CHILD | WS_DISABLED, //dwStyle
		WindowWidth/2 -150/2 -150 -10, //x
		WindowHeight -100 -10, //y
		150, //nWidth
		25, //nHeight
		hWndParent, //hWndParent
		(HMENU)ID_StopButton, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);

	VoiceList = CreateWindow
	(
		WC_COMBOBOX,//_T("COMBOBOX"), //lpClassName
		NULL, //lpWindowName
		CBS_DROPDOWN | WS_VISIBLE | WS_CHILD, //dwStyle
		WindowWidth/2 -150/2, //x
		WindowHeight -100 -35, //y
		150, //nWidth
		25, //nHeight
		hWndParent, //hWndParent
		(HMENU)ID_VoiceList, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);
	//Query Voices
	IEnumSpObjectTokens* EnumerateVoices;
	ISpObjectToken* SpeechToken;
	CSpDynamicString DescriptionString;
	SpEnumTokens(SPCAT_VOICES, NULL, NULL, &EnumerateVoices);
	int Index = 0;
	while (EnumerateVoices->Next(1, &SpeechToken, NULL) == S_OK)
	{
		SpGetDescription(SpeechToken, &DescriptionString);
		SendMessage(VoiceList, CB_ADDSTRING, 0, (LPARAM)DescriptionString.m_psz);
		SendMessage(VoiceList, CB_SETITEMDATA, Index++, (LPARAM)SpeechToken);
	}
	
	VoiceRateSlider = CreateWindow
	(
		TRACKBAR_CLASS, //lpClassName
		NULL, //lpWindowName
		TBS_ENABLESELRANGE | WS_VISIBLE | WS_CHILD, //dwStyle
		WindowWidth/2 -150/2, //x
		WindowHeight -100 -10, //y
		150, //nWidth
		25, //nHeight
		hWndParent, //hWndParent
		(HMENU)ID_VoiceRate, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);
	TextReader->GetRate(&VoiceRate);
	SendMessage(VoiceRateSlider, TBM_SETRANGE, TRUE, MAKELONG(SPMIN_RATE, SPMAX_RATE));
	SendMessage(VoiceRateSlider, TBM_SETPOS, TRUE, VoiceRate);
	SendMessage(VoiceRateSlider, TBM_SETPAGESIZE, TRUE, 5);

	AutoClipboardCheckbox = CreateWindow
	(
		WC_BUTTON, //lpClassName
		_T("Auto-Clipboard"), //lpWindowName
		BS_AUTOCHECKBOX | WS_VISIBLE | WS_CHILD, //dwStyle
		WindowWidth/2 +150/2 +10, //x
		WindowHeight -100 -35, //y
		150, //nWidth
		25, //nHeight
		hWndParent, //hWndParent
		(HMENU)ID_AutoClipboardCheckbox, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);

	AutoReadCheckbox = CreateWindow
	(
		WC_BUTTON, //lpClassName
		_T("Auto-Read"), //lpWindowName
		BS_AUTOCHECKBOX | WS_VISIBLE | WS_CHILD, //dwStyle
		WindowWidth/2 +150/2 +10, //x
		WindowHeight -100 -10, //y
		150, //nWidth
		25, //nHeight
		hWndParent, //hWndParent
		(HMENU)ID_AutoReadCheckbox, //hMenu
		NULL, //hInstance
		NULL //lpParam
	);
}

void InitializeControls()
{
	if (FAILED(SendMessage(VoiceList, CB_SELECTSTRING, (WPARAM)-1, (LPARAM)_T("VW Misaki"))))
	{
		SendMessage(VoiceList, CB_SETCURSEL, (WPARAM)0, 0);
	}
	SetFocus(VoiceList);

	SendMessage(VoiceRateSlider, TBM_SETPOS, TRUE, 1);
	SendMessage(MainWindowHandle, WM_HSCROLL, 0, 0);

	SendMessage(AutoReadCheckbox, BM_CLICK, TRUE, 0);

	SendMessage(AutoClipboardCheckbox, BM_CLICK, TRUE, 0);

	CSpStreamFormat StreamFormat;
	StreamFormat.AssignFormat(SPSF_48kHz16BitStereo);
	ISpAudio* AudioFormat;
	SpCreateDefaultObjectFromCategoryId(SPCAT_AUDIOOUT, &AudioFormat);
	AudioFormat->SetFormat(StreamFormat.FormatId(), StreamFormat.WaveFormatExPtr());
	TextReader->SetOutput(AudioFormat, FALSE);
}

LRESULT WINAPI MainWindowProcedure
	(
		HWND hWnd,
		UINT uMsg,
		WPARAM wParam,
		LPARAM lParam
	)
{
	switch(uMsg)
	{
	//Text Reader Engine Events
	case WM_APP:
		TextReaderEvents();
		break;

	//Clipboard Events
	case WM_CHANGECBCHAIN:
		if ((HWND)wParam == ClipboardViewer)
			ClipboardViewer = (HWND)lParam;
		return 0;
	case WM_DRAWCLIPBOARD:
	{	
		OpenClipboard(NULL);
		int Position = SendMessage(TextBox[AutoClipboardTabIndex], WM_GETTEXTLENGTH, 0, 0);
		SendMessage(TextBox[AutoClipboardTabIndex], EM_SETSEL, Position, Position);
		SendMessage
		(
			TextBox[AutoClipboardTabIndex],
			EM_REPLACESEL,
			0,
			(LPARAM)GetClipboardData(CF_UNICODETEXT)
		);
		CloseClipboard();
		if (ReadState == StoppedState)
		{
			if (SendMessage(AutoReadCheckbox, BM_GETSTATE, 0, 0) == BST_CHECKED)
				SendMessage(MainWindowHandle, WM_COMMAND, ID_ReadPauseResumeButton, 0);
		}
		break;
	}
	//Window Controls Events
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
			case ID_SelectAll:
				if (HIWORD(wParam)==1)
					SendMessage(GetFocus(), EM_SETSEL, (WPARAM)0, (LPARAM)-1);
				break;
			case ID_ReadPauseResumeButton:
				switch (ReadState)
				{
				case StoppedState:
					ReadState = ReadingState;
					SendMessage(ReadPauseResumeButton, WM_SETTEXT, 0, (LPARAM)ReadPauseResumeText[ReadState]);
					EnableWindow(StopButton, TRUE);
					EnableWindow(VoiceList, FALSE);
					EnableWindow(VoiceRateSlider, FALSE);
					TextLength = SendMessage(TextBox[ReadingPageTabIndex], WM_GETTEXTLENGTH, 0, 0);
					ReadText();
					break;
				case ReadingState:
					ReadState = PausedState;
					SendMessage(ReadPauseResumeButton, WM_SETTEXT, 0, (LPARAM)ReadPauseResumeText[ReadState]);
					TextReader->Pause();
					break;
				case PausedState:
					ReadState = ReadingState;
					SendMessage(ReadPauseResumeButton, WM_SETTEXT, 0, (LPARAM)ReadPauseResumeText[ReadState]);
					TextReader->Resume();
					break;
				}
				break;
			case ID_StopButton:
				ReadState = StoppedState;
				SendMessage(ReadPauseResumeButton, WM_SETTEXT, 0, (LPARAM)ReadPauseResumeText[ReadState]);
				SendMessage(TextBox[ReadingPageTabIndex], EM_SETREADONLY, FALSE, 0);
				EnableWindow(StopButton, FALSE);
				EnableWindow(VoiceList, TRUE);
				EnableWindow(VoiceRateSlider, TRUE);
				TextReader->Resume();
				TextReader->Speak(NULL, SPF_PURGEBEFORESPEAK, 0);
				break;
			case ID_VoiceList:
				TextReader->SetVoice
				(
					(ISpObjectToken*)SendMessage
					(
						VoiceList,
						CB_GETITEMDATA,
						SendMessage(VoiceList, CB_GETCURSEL, 0, 0),
						0
					)
				);
				break;
			case ID_AutoClipboardCheckbox:
				if (SendMessage(AutoClipboardCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED)
					ClipboardViewer = SetClipboardViewer(hWnd);
				else
					ChangeClipboardChain(hWnd, ClipboardViewer);
				break;
		}
		return 0;
	case WM_HSCROLL:
		TextReader->SetRate(SendMessage(VoiceRateSlider, TBM_GETPOS, 0, 0));
		break;

	//Tab Controls Events
	case WM_NOTIFY:
		if (SendMessage(TabControls, TCM_GETCURSEL, 0, 0))
		{
			ShowWindow(TextBox[ReadingPageTabIndex], SW_HIDE);
			ShowWindow(TextBox[AutoClipboardTabIndex], SW_SHOW);
		} else
		{
			ShowWindow(TextBox[AutoClipboardTabIndex], SW_HIDE);
			ShowWindow(TextBox[ReadingPageTabIndex], SW_SHOW);
		}
	break;

	//Main Window Events
	case WM_SIZE:
	{
		int LOPARAM = lParam & 0xFFFF;
		int HIPARAM = lParam >> 16;
		SetWindowPos(TabControls, NULL, 0, 0, LOPARAM, HIPARAM-100, SWP_DRAWFRAME);
		SetWindowPos(TextBox[ReadingPageTabIndex], NULL, 5, 30, LOPARAM-10, HIPARAM-100-35, SWP_DRAWFRAME);
		SetWindowPos(TextBox[AutoClipboardTabIndex], NULL, 5, 30, LOPARAM-10, HIPARAM-100-35, SWP_DRAWFRAME);
		SetWindowPos(AutoClipboardCheckbox, NULL, LOPARAM/2 +150/2 +10, HIPARAM-100+10, 150, 25, SWP_DRAWFRAME);
		SetWindowPos(AutoReadCheckbox, NULL, LOPARAM/2 +150/2 +10, HIPARAM-100+35, 150, 25, SWP_DRAWFRAME);
		SetWindowPos(ReadPauseResumeButton, NULL, LOPARAM/2 -150/2 -150 -10, HIPARAM-100+10, 150, 25, SWP_DRAWFRAME);
		SetWindowPos(StopButton, NULL, LOPARAM/2 -150/2 -150 -10, HIPARAM-100+35, 150, 25, SWP_DRAWFRAME);
		SetWindowPos(VoiceList, NULL, LOPARAM/2 -150/2, HIPARAM-100+10, 150, 25, SWP_DRAWFRAME);
		SetWindowPos(VoiceRateSlider, NULL, LOPARAM/2 -150/2, HIPARAM-100+35, 150, 25, SWP_DRAWFRAME);
		return 0;
	}
	case WM_CREATE:
		CreateControls(hWnd);
		return 0;
	case WM_DESTROY:
		ChangeClipboardChain(hWnd, ClipboardViewer);
		DestroyAcceleratorTable(SelectAll);
		PostQuitMessage(EXIT_SUCCESS);
		return 0;
	}
	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

void CreateMainWindow()
{
	WNDCLASSEX WindowClassStructure =
	{
		sizeof(WNDCLASSEX),					//cbSize
		CS_HREDRAW | CS_VREDRAW,			//style
		MainWindowProcedure,				//lpfnWndProc
		0,									//cbClsExtra
		0,									//cbWndExtra
		NULL,								//hInstance
		LoadIcon(NULL, IDI_INFORMATION),	//hIcon
		LoadCursor(NULL, IDC_ARROW),		//hCursor
		(HBRUSH)(COLOR_BTNFACE + 1),		//hbrBackground
		NULL,								//lpszMenuName
		_T("MainWindow"),					//lpszClassName
		LoadIcon(NULL, IDI_INFORMATION)		//hIconSm
	};
	RegisterClassEx(&WindowClassStructure);

	MainWindowHandle = CreateWindow
	(
		_T("MainWindow"),					//lpClassName
		_T("Clipboard Reader"),				//lpWindowName
		WS_VISIBLE | WS_OVERLAPPEDWINDOW,	//dwStyle
		CW_USEDEFAULT,						//x
		CW_USEDEFAULT,						//y
		WindowWidth,						//nWidth
		WindowHeight,						//nHeight
		NULL,								//hWndParent
		NULL,								//hMenu
		NULL,								//hInstance
		NULL								//lpParam
	);
}

WPARAM MessageLoop()
{
	MSG MessageStructure = { 0 };
	while (MessageStructure.message != WM_QUIT)
	{
		if (GetMessage(&MessageStructure, NULL, 0, 0))
		{
			if (!TranslateAccelerator(MainWindowHandle, SelectAll, &MessageStructure))
			{
				TranslateMessage(&MessageStructure);
				DispatchMessage(&MessageStructure);
			}
		}
	}
	return MessageStructure.wParam;
}

int APIENTRY _tWinMain
	(
		HINSTANCE hInstance,
		HINSTANCE hPrevInstance,
		LPTSTR lpCmdLine,
		int nCmdShow
	)
{
	CoInitialize(NULL);
	CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&TextReader);
	TextReader->SetInterest(SPFEI_ALL_TTS_EVENTS, SPFEI_ALL_TTS_EVENTS);

	CreateMainWindow();
	InitializeControls();

	WPARAM MessageResponse = MessageLoop();
	
	TextReader->Release();
	CoUninitialize();

	return MessageResponse;
}