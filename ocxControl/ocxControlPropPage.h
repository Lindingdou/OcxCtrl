#pragma once

// ocxControlPropPage.h: CocxControlPropPage 属性页类的声明。


// CocxControlPropPage : 请参阅 ocxControlPropPage.cpp 了解实现。

class CocxControlPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CocxControlPropPage)
	DECLARE_OLECREATE_EX(CocxControlPropPage)

// 构造函数
public:
	CocxControlPropPage();

// 对话框数据
	enum { IDD = IDD_PROPPAGE_OCXCONTROL };

// 实现
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 消息映射
protected:
	DECLARE_MESSAGE_MAP()
};

