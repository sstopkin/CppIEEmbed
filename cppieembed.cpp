#define STRICT
#define NOCRYPT

#include "webform.h"
#include "webwindow.h"
#include "webformdispatchimpl.h"
#include "jsobject.h"
#include <windows.h>
#include <tchar.h>
#include <map>
#include <string>
#include <sstream>
#include <cctype>
#include <algorithm>

bool loaded;
WebWindow* webWindow;
JSObject *jsobj;
WebformDispatchImpl *webformDispatchImpl;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	OleInitialize(0);

	jsobj = new JSObject();
	jsobj->AddRef();
	webformDispatchImpl = new WebformDispatchImpl(jsobj);

	webWindow = new WebWindow(webformDispatchImpl);
	webWindow->Create(hInstance, 200, 100, 600, 690, true);

	TCHAR Buffer[MAX_PATH];
	DWORD dwRet;
	dwRet = GetCurrentDirectory(MAX_PATH, Buffer);
	std::string appPath = Buffer;
	TCHAR* indexPath = _T("../html/index.html");
	appPath = appPath + indexPath;

	webWindow->webForm->Go(appPath.c_str());

	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	delete webWindow;
	delete webformDispatchImpl;
	jsobj->Release();

	OleUninitialize();
	return (int)msg.wParam;
}
