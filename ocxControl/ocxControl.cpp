// ocxControl.cpp: CocxControlApp 和 DLL 注册的实现。

#include "pch.h"
#include "framework.h"
#include "ocxControl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CocxControlApp theApp;

const GUID CDECL _tlid = {0xeb5cd14b,0xe702,0x4de0,{0x98,0x07,0x00,0x38,0x46,0x13,0x08,0x8b}};
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;



// CocxControlApp::InitInstance - DLL 初始化

BOOL CocxControlApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO:  在此添加您自己的模块初始化代码。
	}

	return bInit;
}



// CocxControlApp::ExitInstance - DLL 终止

int CocxControlApp::ExitInstance()
{
	// TODO:  在此添加您自己的模块终止代码。

	return COleControlModule::ExitInstance();
}



// DllRegisterServer - 将项添加到系统注册表

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}



// DllUnregisterServer - 将项从系统注册表中移除

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}
