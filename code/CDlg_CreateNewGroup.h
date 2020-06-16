#pragma once


// CDlg_CreateNewGroup 对话框

class CDlg_CreateNewGroup : public CDialogEx
{
	DECLARE_DYNAMIC(CDlg_CreateNewGroup)

public:
	CDlg_CreateNewGroup(int q , CStringArray** qN,int* f , CString g, CWnd* pParent = nullptr);   // 标准构造函数
	virtual ~CDlg_CreateNewGroup();
	int qNum;					//传参过来的题目数，用于根据题目数生成勾选框
	CStringArray* qNames;		//存传过来的各列表头（各个题目的名字）
	int* groupFlag;				//传参过来的标志数组，用于记录这个模块包含哪些题目
	CString groupName;			//这个正在编辑的题组的名字
// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DIALOG_CREATE_NEW_GROUP};
#endif

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
private:
	CStatic v_tips_text;
	int* bntID;//存所有bnt控件的id
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnClose();
	afx_msg void OnBnClickedCancel();
	afx_msg int CountFlag();//统计各个勾选框的选中情况
};
