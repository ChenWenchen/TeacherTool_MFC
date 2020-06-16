#pragma once


// CTeacherToolDlg 对话框
class CTeacherToolDlg : public CDialogEx
{
// 构造
public:
	CTeacherToolDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_TEACHERTOOL_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	BOOL v_ifRemoveFirst;	//是否去掉姓氏
	CListCtrl v_reportTable;//学生成绩展示报表

	afx_msg void OnBnClickedImportStu();//导入学生成绩
	afx_msg void OnBnClickedImportSpk();//导入话术列表

	CEdit v_newGroupName;//添加新题组的名字
	afx_msg void OnBnClickedAddGroupBnt();//点击添加题组、模块
	
	
	CComboBox v_groupList;//已有的题组/模块下拉框
	afx_msg void OnCbnSelchangeCombo2();//切换选中的题组模块
	afx_msg void OnBnClickedCountGroupBnt();//按模块统计

	afx_msg void OnBnClickedGenerateReport();//生成报告 按钮
	
	virtual void OnOK();
	//afx_msg void OnBnClickedRemoveFirst();
	afx_msg void OnBnClickedExportExcel();
	afx_msg void OnBnClickedExportTxt();

	afx_msg void CTeacherToolDlg::doCalculate();
	afx_msg void CTeacherToolDlg::showReport();
	afx_msg void CTeacherToolDlg::clearTable();
	afx_msg void OnBnClickedLinkWeb();
	afx_msg void SetAddGroupVisble(BOOL ifVisble);
};
