// ocxControlPropPage.cpp : CocxControlPropPage 属性页类的实现。

#include "pch.h"
#include "framework.h"
#include "ocxControl.h"
#include "ocxControlPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CocxControlPropPage, COlePropertyPage)

// 消息映射

BEGIN_MESSAGE_MAP(CocxControlPropPage, COlePropertyPage)
END_MESSAGE_MAP()

// 初始化类工厂和 guid

IMPLEMENT_OLECREATE_EX(CocxControlPropPage, "MFCACTIVEXCONT.ocxControlPropPage.1",
	0x705a2d0c,0x72bb,0x4f09,0xbf,0xb9,0xfc,0x17,0x3f,0xcf,0x1b,0x3a)

// CocxControlPropPage::CocxControlPropPageFactory::UpdateRegistry -
// 添加或移除 CocxControlPropPage 的系统注册表项

BOOL CocxControlPropPage::CocxControlPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_OCXCONTROL_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, nullptr);
}

// CocxControlPropPage::CocxControlPropPage - 构造函数

CocxControlPropPage::CocxControlPropPage() :
	COlePropertyPage(IDD, IDS_OCXCONTROL_PPG_CAPTION)
{
}

// CocxControlPropPage::DoDataExchange - 在页和属性间移动数据

void CocxControlPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_PostProcessing(pDX);
}

// CocxControlPropPage 消息处理程序
