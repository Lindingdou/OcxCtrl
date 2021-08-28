// ocxControlCtrl.cpp : CocxControlCtrl ActiveX 控件类的实现。

#include "pch.h"
#include "framework.h"
#include "ocxControl.h"
#include "ocxControlCtrl.h"
#include "ocxControlPropPage.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CocxControlCtrl, COleControl)

// 消息映射

BEGIN_MESSAGE_MAP(CocxControlCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()

// 调度映射

BEGIN_DISPATCH_MAP(CocxControlCtrl, COleControl)
	DISP_FUNCTION_ID(CocxControlCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

// 事件映射

BEGIN_EVENT_MAP(CocxControlCtrl, COleControl)
END_EVENT_MAP()

// 属性页

// TODO: 根据需要添加更多属性页。请记住增加计数!
BEGIN_PROPPAGEIDS(CocxControlCtrl, 1)
	PROPPAGEID(CocxControlPropPage::guid)
END_PROPPAGEIDS(CocxControlCtrl)

// 初始化类工厂和 guid

IMPLEMENT_OLECREATE_EX(CocxControlCtrl, "MFCACTIVEXCONTRO.ocxControlCtrl.1",
	0xf6f55a23,0x02a6,0x4798,0xaf,0x71,0xa8,0x3b,0xc2,0xa5,0xba,0x30)

// 键入库 ID 和版本

IMPLEMENT_OLETYPELIB(CocxControlCtrl, _tlid, _wVerMajor, _wVerMinor)

// 接口 ID

const IID IID_DocxControl = {0xc9608f5a,0x6c84,0x4e33,{0x86,0xa2,0x10,0xa0,0x33,0xfe,0xa5,0xb4}};
const IID IID_DocxControlEvents = {0xf868096b,0xd85a,0x4ad8,{0xbd,0x69,0x15,0x23,0x52,0x3e,0x4c,0xe6}};

// 控件类型信息

static const DWORD _dwocxControlOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CocxControlCtrl, IDS_OCXCONTROL, _dwocxControlOleMisc)

// CocxControlCtrl::CocxControlCtrlFactory::UpdateRegistry -
// 添加或移除 CocxControlCtrl 的系统注册表项

BOOL CocxControlCtrl::CocxControlCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO:  验证您的控件是否符合单元模型线程处理规则。
	// 有关更多信息，请参考 MFC 技术说明 64。
	// 如果您的控件不符合单元模型规则，则
	// 必须修改如下代码，将第六个参数从
	// afxRegApartmentThreading 改为 0。

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_OCXCONTROL,
			IDB_OCXCONTROL,
			afxRegApartmentThreading,
			_dwocxControlOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


// CocxControlCtrl::CocxControlCtrl - 构造函数

CocxControlCtrl::CocxControlCtrl()
{
	InitializeIIDs(&IID_DocxControl, &IID_DocxControlEvents);
	// TODO:  在此初始化控件的实例数据。
}

// CocxControlCtrl::~CocxControlCtrl - 析构函数

CocxControlCtrl::~CocxControlCtrl()
{
	// TODO:  在此清理控件的实例数据。
}

// CocxControlCtrl::OnDraw - 绘图函数

void CocxControlCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /* rcInvalid */)
{
	if (!pdc)
		return;

	// TODO:  用您自己的绘图代码替换下面的代码。
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	pdc->Ellipse(rcBounds);
}

// CocxControlCtrl::DoPropExchange - 持久性支持

void CocxControlCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: 为每个持久的自定义属性调用 PX_ 函数。
}


// CocxControlCtrl::OnResetState - 将控件重置为默认状态

void CocxControlCtrl::OnResetState()
{
	COleControl::OnResetState();  // 重置 DoPropExchange 中找到的默认值

	// TODO:  在此重置任意其他控件状态。
}


// CocxControlCtrl::AboutBox - 向用户显示“关于”框

void CocxControlCtrl::AboutBox()
{
	CDialogEx dlgAbout(IDD_ABOUTBOX_OCXCONTROL);
	dlgAbout.DoModal();
}


// CocxControlCtrl 消息处理程序
