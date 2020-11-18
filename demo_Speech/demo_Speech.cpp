// demo_Speech.cpp : �������̨Ӧ�ó������ڵ㡣
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
//�����µ��׽���֮ǰ��Ҫ����һ������Ws2_32.dll��ĺ���,����������Ϳͻ������Ӳ���
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

TCHAR szAppName[] = TEXT("��������Demo");
BOOL b_initSR;
BOOL b_Cmd_Grammar;
CComPtr<ISpRecoContext> m_cpRecoCtxt;//����ʶ�����ӿ�
CComPtr<ISpRecoGrammar> m_cpCmdGramma;//ʶ���﷨
CComPtr<ISpRecognizer> m_cpRecoEngine; //����ʶ������
CComPtr<ISpRecoGrammar> m_cpDictationGrammar;//ʶ���﷨ ��˵ʽ
const wchar_t *configFile = L"C:\\Users\\share\\Desktop\\demo_Speech\\x64\\Debug\\cmd.xml";
static int UDP_PORT = 7001;
WSADATA wsaData;                                    //ָ��WinSocket��Ϣ�ṹ��ָ��
SOCKET sockListener;                                //�����׽���
SOCKADDR_IN saUdpServ;                              //ָ��ͨ�Ŷ���Ľṹ��ָ��                          
BOOL fBroadcast = TRUE;
int speak(wchar_t *str);
void MSSListen(HWND hwnd);

int WINAPI main(HINSTANCE hInstance, HINSTANCE hPrevInstance, PSTR szCmdLine, int iCmdShow)
{
	HWND hwnd;
	MSG msg;
	WNDCLASS wndclass;

	//������ṹ���ʼ��ֵ
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

	//ע�ᴰ����
	if (!RegisterClass(&wndclass))
	{
		//ʧ�ܺ���ʾ������
		MessageBox(NULL, TEXT("This program requires Windows NT!"), szAppName, MB_ICONERROR);
		return 0;
	}

	//��������
	hwnd = CreateWindow(szAppName,
		TEXT("����ʶ��"),
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
	if (WSAStartup(MAKEWORD(1, 1), &wsaData) != 0)     //����WinSocket�ĳ�ʼ��
	{
		printf("Can't initiates windows socket!Program stop.\n");//��ʼ��ʧ�ܷ���-1
		return -1;
	}
	sockListener = socket(PF_INET, SOCK_DGRAM, 0);
	//       setsockopt�������������׽ӿ�ѡ��
	//       ���ù㲥��ʽ�뽫��������������ΪSO_BROADCAST
	setsockopt(sockListener, SOL_SOCKET, SO_BROADCAST, (CHAR *)&fBroadcast, sizeof(BOOL));
	//  �������ã�ע��Ҫ��IP��ַ��ΪINADDR_BROADCAST����ʾ���͹㲥UDP���ݱ�
	saUdpServ.sin_family = AF_INET;
	saUdpServ.sin_addr.s_addr = htonl(INADDR_BROADCAST);
	saUdpServ.sin_port = htons(UDP_PORT);


	//��ʾ����
	//ShowWindow(hwnd, iCmdShow);
	UpdateWindow(hwnd);
	MSSListen(hwnd);
	//������Ϣѭ��
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);//������Ϣ
		DispatchMessage(&msg);//�ַ���Ϣ
	}
	return msg.wParam;
}

/*
*��Ϣ�ص�����,�ɲ���ϵͳ����
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
		//��ʼ��COM�˿�
		::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
		//����ʶ�������Ľӿ�
		HRESULT hr = m_cpRecoEngine.CoCreateInstance(CLSID_SpSharedRecognizer);
		if (SUCCEEDED(hr))
		{
			hr = m_cpRecoEngine->CreateRecoContext(&m_cpRecoCtxt);
		}
		else
		{
			MessageBox(hwnd, TEXT("����ʵ��������"), TEXT("��ʾ"), S_OK);
		}
		//����ʶ����Ϣ��ʹ�����ʱ�̼���������Ϣ
		if (SUCCEEDED(hr))
		{
			hr = m_cpRecoCtxt->SetNotifyWindowMessage(hwnd, WM_RECOEVENT, 0, 0);
		}
		else
		{
			MessageBox(hwnd, TEXT("���������Ľӿڳ���"), TEXT("��ʾ"), S_OK);
		}
		//�������Ǹ���Ȥ���¼�
		if (SUCCEEDED(hr))
		{
			ULONGLONG ullMyEvents = SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_RECOGNITION) | SPFEI(SPEI_SOUND_END);
			hr = m_cpRecoCtxt->SetInterest(ullMyEvents, ullMyEvents);
		}
		else
		{
			MessageBox(hwnd, TEXT("����ʶ����Ϣ����"), TEXT("��ʾ"), S_OK);
		}
		//�����﷨����
		b_Cmd_Grammar = TRUE;
		if (FAILED(hr))
		{
			MessageBox(hwnd, TEXT("�����﷨�������"), TEXT("��ʾ"), S_OK);
		}
		hr = m_cpRecoCtxt->CreateGrammar(GID_CMD_GR, &m_cpCmdGramma);
		hr = m_cpCmdGramma->LoadCmdFromFile(configFile, SPLO_DYNAMIC);
		if (FAILED(hr))
		{
			MessageBox(hwnd, TEXT("�����ļ��򿪳���"), TEXT("��ʾ"), S_OK);
		}
		b_initSR = TRUE;
		//�ڿ�ʼʶ��ʱ�������﷨����ʶ��
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
				//ȡ����Ϣ���

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
					//commandΪʶ������ַ���
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

#pragma comment(lib,"ole32.lib")//CoInitialize CoCreateInstance ��Ҫ����ole32.dll

/*
*�����ϳɺ������ʶ��ַ���str
*/
int speak(wchar_t *str)
{
	ISpVoice * pVoice = NULL;
	::CoInitialize(NULL);
	//���ISpVoice�ӿ�
	long hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void **)&pVoice);
	hr = pVoice->Speak(str, 0, NULL);
	pVoice->Release();
	pVoice = NULL;
	//ǧ��Ҫ����
	::CoUninitialize();
	return TRUE;
}

void MSSListen(HWND hwnd)
{
	//��ʼ��COM�˿�
	::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	//����ʶ�������Ľӿ�
	HRESULT hr = m_cpRecoEngine.CoCreateInstance(CLSID_SpSharedRecognizer);
	if (SUCCEEDED(hr))
	{
		hr = m_cpRecoEngine->CreateRecoContext(&m_cpRecoCtxt);
	}
	else
	{
		MessageBox(hwnd, TEXT("����ʵ��������"), TEXT("��ʾ"), S_OK);
	}
	//����ʶ����Ϣ��ʹ�����ʱ�̼���������Ϣ
	if (SUCCEEDED(hr))
	{
		hr = m_cpRecoCtxt->SetNotifyWindowMessage(hwnd, WM_RECOEVENT, 0, 0);
	}
	else
	{
		MessageBox(hwnd, TEXT("���������Ľӿڳ���"), TEXT("��ʾ"), S_OK);
	}
	//�������Ǹ���Ȥ���¼�
	if (SUCCEEDED(hr))
	{
		ULONGLONG ullMyEvents = SPFEI(SPEI_SOUND_START) | SPFEI(SPEI_RECOGNITION) | SPFEI(SPEI_SOUND_END);
		hr = m_cpRecoCtxt->SetInterest(ullMyEvents, ullMyEvents);
	}
	/*else
	{
	MessageBox(hwnd, TEXT("����ʶ����Ϣ����"), TEXT("��ʾ"), S_OK);
	}*/
	//�����﷨����

	//hr = m_cpRecoCtxt->CreateGrammar(3333, &m_cpDictationGrammar);
	//if (SUCCEEDED(hr))
	//{
	//  hr = m_cpDictationGrammar->LoadDictation(NULL, SPLO_STATIC);//���شʵ�
	//}
	b_Cmd_Grammar = TRUE;
	if (FAILED(hr))
	{
		MessageBox(hwnd, TEXT("�����﷨�������"), TEXT("��ʾ"), S_OK);
	}
	hr = m_cpRecoCtxt->CreateGrammar(GID_CMD_GR, &m_cpCmdGramma);
	hr = m_cpCmdGramma->LoadCmdFromFile(configFile, SPLO_DYNAMIC);
	if (FAILED(hr))
	{
		MessageBox(hwnd, TEXT("�����ļ��򿪳���"), TEXT("��ʾ"), S_OK);
	}
	b_initSR = TRUE;
	//�ڿ�ʼʶ��ʱ�������﷨����ʶ��
	//hr = m_cpDictationGrammar->SetDictationState(SPRS_ACTIVE);//dictation
	hr = m_cpCmdGramma->SetRuleState(NULL, NULL, SPRS_ACTIVE);
	hr = m_cpRecoEngine->SetRecoState(SPRST_ACTIVE);
	m_cpRecoCtxt->WaitForNotifyEvent(1);
	::CoUninitialize();
}
