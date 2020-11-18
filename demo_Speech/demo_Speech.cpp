// demo_Speech.cpp : 定义控制台应用程序的入口点。
//
#include "stdafx.h"
#include <Windows.h>
#include <atlstr.h>
#include <sphelper.h>
#include <sapi.h>
#include <comutil.h>
#include <string.h>
#include <iostream>
#include<algorithm>
#include <vector>
#include <winsock.h>
using namespace std;
//创建新的套接字之前需要调用一个引入Ws2_32.dll库的函数,否则服务器和客户端连接不上
#pragma comment(lib,"ws2_32.lib")
#pragma comment(lib,"sapi.lib")
#ifdef _UNICODE
#pragma   comment(lib,   "comsuppw.lib")  //_com_util::ConvertBSTRToString
#else
#pragma   comment(lib,   "comsupp.lib")  //_com_util::ConvertBSTRToString
#endif

#define GID_CMD_GR 333333
#define WM_RECOEVENT WM_USER+1

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

TCHAR szAppName[] = TEXT("语音控制Demo");
BOOL b_initSR;
BOOL b_Cmd_Grammar;
CComPtr<ISpRecoContext> m_cpRecoCtxt;//语音识别程序接口
CComPtr<ISpRecoGrammar> m_cpCmdGramma;//识别语法
CComPtr<ISpRecognizer> m_cpRecoEngine; //语音识别引擎
CComPtr<ISpRecoGrammar> m_cpDictationGrammar;//识别语法 听说式
const wchar_t *configFile = L"C:\\Users\\share\\Desktop\\demo_Speech\\x64\\Debug\\cmd.xml";
static int UDP_PORT = 7001;
WSADATA wsaData;                                    //指向WinSocket信息结构的指针
SOCKET sockListener;                                //创建套接字
SOCKADDR_IN saUdpServ;                              //指向通信对象的结构体指针                          
BOOL fBroadcast = TRUE;
int speak(wchar_t *str);
void MSSListen(HWND hwnd);

int WINAPI main(HINSTANCE hInstance, HINSTANCE hPrevInstance, PSTR szCmdLine, int iCmdShow)
{
	HWND hwnd;
	MSG msg;
	WNDCLASS wndclass;

	//窗口类结构体初始化值
	wndclass.cbClsExtra = 0;
	wndclass.cbWndExtra = 0;
	wndclass.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
	wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
	wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wndclass.hInstance = hInstance;
	wndclass.lpfnWndProc = WndProc;
	wndclass.lpszClassName = szAppName;
	wndclass.lpszMenuName = NULL;
	wndclass.style = CS_HREDRAW | CS_VREDRAW;

	//注册窗口类
	if (!RegisterClass(&wndclass))
	{
		//失败后提示并返回
		MessageBox(NULL, TEXT("This program requires Windows NT!"), szAppName, MB_ICONERROR);
		return 0;
	}

	//创建窗口
	hwnd = CreateWindow(szAppName,
		TEXT("语音识别"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		NULL,
		NULL,
		hInstance,
		NULL);


	//socket
	if (WSAStartup(MAKEWORD(1, 1), &wsaData) != 0)     //进行WinSocket的初始化
	{
		printf("Can't initiates windows socket!Program stop.\n");//初始化失败返回-1
		return -1;
	}
	sockListener = socket(PF_INET, SOCK_DGRAM, 0);
	//       setsockopt函数用于设置套接口选项
	//       采用广播形式须将第三个参数设置为SO_BROADCAST
	setsockopt(sockListener, SOL_SOCKET, SO_BROADCAST, (CHAR *)&fBroadcast, sizeof(BOOL));
	//  参数设置，注意要将IP地址设为INADDR_BROADCAST，表示发送广播UDP数据报
	saUdpServ.sin_family = AF_INET;
	saUdpServ.sin_addr.s_addr = htonl(INADDR_BROADCAST);
	saUdpServ.sin_port = htons(UDP_PORT);


	//显示窗口
	//ShowWindow(hwnd, iCmdShow);
	UpdateWindow(hwnd);
	MSSListen(hwnd);
	//进入消息循环
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);//翻译消息
		DispatchMessage(&msg);//分发消息
	}
	return msg.wParam;
}

/*
*消息回调函数,由操作系统调用
*/
LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	HDC hdc;
	PAINTSTRUCT ps;

	switch (message)
	{
	case WM_CREATE:
	{
		break;
		//初始化COM端口
		::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
		//创建识别上下文接口
		HRESULT hr = m_cpRecoEngine.CoCreateInstance(CLSID_SpSharedRecognizer);
		if (SUCCEEDED(hr))
		{
			hr = m_cpRecoEngine->CreateRecoContext(&m_cpRecoCtxt);
		}
		else
		{
			MessageBox(hwnd, TEXT("引擎实例化出错"), TEXT("提示"), S_OK);
		}
		//设置识别消息，使计算机时刻监听语音消息
		if (SUCCEEDED(hr))
		{
			hr = m_cpRecoCtxt->SetNotifyWindowMessage(hwnd, WM_RECOEVENT, 0, 0);
		}
		else
		{
			MessageBox(hwnd, TEXT("创建上下文接口出错"), TEXT("提示"), S_OK);
		}
		//设置我们感兴趣的事件
		if (SUCCEEDED(hr))
		{
			ULONGLONG ullMyEvents = SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_RECOGNITION) | SPFEI(SPEI_SOUND_END);
			hr = m_cpRecoCtxt->SetInterest(ullMyEvents, ullMyEvents);
		}
		else
		{
			MessageBox(hwnd, TEXT("设置识别消息出错"), TEXT("提示"), S_OK);
		}
		//创建语法规则
		b_Cmd_Grammar = TRUE;
		if (FAILED(hr))
		{
			MessageBox(hwnd, TEXT("创建语法规则出错"), TEXT("提示"), S_OK);
		}
		hr = m_cpRecoCtxt->CreateGrammar(GID_CMD_GR, &m_cpCmdGramma);
		hr = m_cpCmdGramma->LoadCmdFromFile(configFile, SPLO_DYNAMIC);
		if (FAILED(hr))
		{
			MessageBox(hwnd, TEXT("配置文件打开出错"), TEXT("提示"), S_OK);
		}
		b_initSR = TRUE;
		//在开始识别时，激活语法进行识别
		hr = m_cpCmdGramma->SetRuleState(NULL, NULL, SPRS_ACTIVE);
		break;
	}
	case WM_RECOEVENT:
	{
		RECT rect;
		GetClientRect(hwnd, &rect);
		hdc = GetDC(hwnd);
		USES_CONVERSION;
		CSpEvent event;
		while (event.GetFrom(m_cpRecoCtxt) == S_OK)
		{
			switch (event.eEventId)
			{
			case SPEI_RECOGNITION:
			{
				static const WCHAR wszUnrecognized[] = L"<Unrecognized>";
				CSpDynamicString dstrText;
				//取得消息结果

				if (FAILED(event.RecoResult()->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, TRUE, &dstrText, NULL)))
				{
					dstrText = wszUnrecognized;

				}
				BSTR SRout;
				dstrText.CopyToBSTR(&SRout);
				char * lpszText2 = _com_util::ConvertBSTRToString(SRout);

				if (b_Cmd_Grammar)
				{
					wchar_t* str = new wchar_t();
					MultiByteToWideChar(CP_ACP, NULL, lpszText2, -1, str, 1024);
					//command为识别出的字符串
					std::string command = lpszText2;
					speak((wchar_t*)str);
				}
			}
		}
	}
		break;
	}
	case WM_PAINT:
		hdc = BeginPaint(hwnd, &ps);
		EndPaint(hwnd, &ps);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	}
	return DefWindowProc(hwnd, message, wParam, lParam);
}

#pragma comment(lib,"ole32.lib")//CoInitialize CoCreateInstance 需要调用ole32.dll

/*
*语音合成函数，朗读字符串str
*/
int speak(wchar_t *str)
{
	ISpVoice * pVoice = NULL;
	::CoInitialize(NULL);
	//获得ISpVoice接口
	long hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void **)&pVoice);
	hr = pVoice->Speak(str, 0, NULL);
	pVoice->Release();
	pVoice = NULL;
	//千万不要忘记
	::CoUninitialize();
	return TRUE;
}

void MSSListen(HWND hwnd)
{
	//初始化COM端口
	::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	//创建识别上下文接口
	HRESULT hr = m_cpRecoEngine.CoCreateInstance(CLSID_SpSharedRecognizer);
	if (SUCCEEDED(hr))
	{
		hr = m_cpRecoEngine->CreateRecoContext(&m_cpRecoCtxt);
	}
	else
	{
		MessageBox(hwnd, TEXT("引擎实例化出错"), TEXT("提示"), S_OK);
	}
	//设置识别消息，使计算机时刻监听语音消息
	if (SUCCEEDED(hr))
	{
		hr = m_cpRecoCtxt->SetNotifyWindowMessage(hwnd, WM_RECOEVENT, 0, 0);
	}
	else
	{
		MessageBox(hwnd, TEXT("创建上下文接口出错"), TEXT("提示"), S_OK);
	}
	//设置我们感兴趣的事件
	if (SUCCEEDED(hr))
	{
		ULONGLONG ullMyEvents = SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_RECOGNITION) | SPFEI(SPEI_SOUND_END);
		hr = m_cpRecoCtxt->SetInterest(ullMyEvents, ullMyEvents);
	}
	/*else
	{
	MessageBox(hwnd, TEXT("设置识别消息出错"), TEXT("提示"), S_OK);
	}*/
	//创建语法规则

	//hr = m_cpRecoCtxt->CreateGrammar(3333, &m_cpDictationGrammar);
	//if (SUCCEEDED(hr))
	//{
	//  hr = m_cpDictationGrammar->LoadDictation(NULL, SPLO_STATIC);//加载词典
	//}
	b_Cmd_Grammar = TRUE;
	if (FAILED(hr))
	{
		MessageBox(hwnd, TEXT("创建语法规则出错"), TEXT("提示"), S_OK);
	}
	hr = m_cpRecoCtxt->CreateGrammar(GID_CMD_GR, &m_cpCmdGramma);
	hr = m_cpCmdGramma->LoadCmdFromFile(configFile, SPLO_DYNAMIC);
	if (FAILED(hr))
	{
		MessageBox(hwnd, TEXT("配置文件打开出错"), TEXT("提示"), S_OK);
	}
	b_initSR = TRUE;
	//在开始识别时，激活语法进行识别
	//hr = m_cpDictationGrammar->SetDictationState(SPRS_ACTIVE);//dictation
	hr = m_cpCmdGramma->SetRuleState(NULL, NULL, SPRS_ACTIVE);
	hr = m_cpRecoEngine->SetRecoState(SPRST_ACTIVE);
	m_cpRecoCtxt->WaitForNotifyEvent(1);
	::CoUninitialize();
}
