#pragma once

// ocxControlCtrl.h : CocxControlCtrl ActiveX 控件类的声明。


// CocxControlCtrl : 请参阅 ocxControlCtrl.cpp 了解实现。

class CocxControlCtrl : public COleControl
{
	DECLARE_DYNCREATE(CocxControlCtrl)

// 构造函数
public:
	CocxControlCtrl();

// 重写
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// 实现
protected:
	~CocxControlCtrl();

	DECLARE_OLECREATE_EX(CocxControlCtrl)    // 类工厂和 guid
	DECLARE_OLETYPELIB(CocxControlCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CocxControlCtrl)     // 属性页 ID
	DECLARE_OLECTLTYPE(CocxControlCtrl)		// 类型名称和杂项状态

// 消息映射
	DECLARE_MESSAGE_MAP()

// 调度映射
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// 事件映射
	DECLARE_EVENT_MAP()

// 调度和事件 ID
public:
	enum {
	};
};

