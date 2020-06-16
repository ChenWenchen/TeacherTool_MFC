// CDlg_CreateNewGroup.cpp: 实现文件
//

#include "pch.h"
#include "TeacherTool.h"
#include "CDlg_CreateNewGroup.h"
#include "afxdialogex.h"


// CDlg_CreateNewGroup 对话框

IMPLEMENT_DYNAMIC(CDlg_CreateNewGroup, CDialogEx)

CDlg_CreateNewGroup::CDlg_CreateNewGroup(int q , 
										CStringArray** qN,
										int* f,
										CString g,
										CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_DIALOG_CREATE_NEW_GROUP, pParent)
{
	qNum = q;
	qNames = *qN;
	groupFlag = f;
	groupName = g;

	bntID = new int[q];//存所有bnt控件的id
}

CDlg_CreateNewGroup::~CDlg_CreateNewGroup()
{
}

void CDlg_CreateNewGroup::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TIPS_TEXT, v_tips_text);
}


BEGIN_MESSAGE_MAP(CDlg_CreateNewGroup, CDialogEx)
	ON_BN_CLICKED(IDOK, &CDlg_CreateNewGroup::OnBnClickedOk)
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDCANCEL, &CDlg_CreateNewGroup::OnBnClickedCancel)
END_MESSAGE_MAP()


// CDlg_CreateNewGroup 消息处理程序

// 生成并显示勾选框窗口
BOOL CDlg_CreateNewGroup::OnInitDialog(){
	CDialogEx::OnInitDialog();

	v_tips_text.SetWindowTextW(TEXT("请选择哪些项目属于\""+groupName+"\""));
	SetWindowText(groupName);

	CString tmp;
	int x, y, row, col;
	int COL = 10;					//每行显示几个勾选框
	int ROW = qNum / COL + 1;		//计算几行才够显示所有勾选框
	int k = 0;
	int thisId;						//生成每个bnt的id，注意bnt间不要重复
	for (row = 0; row < ROW; row++) {
		y = 50 + 40 * row;			//计算勾选框的横纵坐标
		for (col = 0; col < COL; col++) {
			if (k < qNum) {
				x = 50 + 80 * col;
				CRect rect;// not (x,y,50,50); !!!!
				rect.left = x;
				rect.top = y;
				rect.right = rect.left + 80;//这个left和top决定每个勾选框的有效点击范围，不要和x，y冲突
				rect.bottom = rect.top + 50;
				
				CButton* MyChk = new CButton();
				tmp = qNames->GetAt(k+1);   //表头[0]是姓名，不要
				thisId = 1234 + k;			//从1234开始，不重复生成bnt的id
				MyChk->Create(tmp, WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, rect, this, thisId);
				bntID[k] = thisId;
				
				k++;
			}
		}
	}

	return TRUE;  // return TRUE unless you set the focus to a control
				  // 异常: OCX 属性页应返回 FALSE
}

//统计各个勾选框的选中情况
int CDlg_CreateNewGroup::CountFlag() {
	int flag = 0;
	int state = -1;
	for (int i = 0; i < qNum; i++) {
		state = ((CButton*)GetDlgItem(1234 + i))->GetCheck();
		groupFlag[i] = state;

		if (state == 1) {
			flag = 1;//有一个勾选框被勾了就ok否则报错
		}
	}
	return flag;
}

void CDlg_CreateNewGroup::OnBnClickedOk(){

	if (CountFlag() == 1) {
		CDialogEx::OnOK();
	}else {
		MessageBox(TEXT("请至少勾选一个项目吧？"));
	}
}

void CDlg_CreateNewGroup::OnClose(){

	if (CountFlag() == 1) {
		CDialogEx::OnClose();
	}else {
		MessageBox(TEXT("请至少勾选一个项目吧？"));
	}
}

void CDlg_CreateNewGroup::OnBnClickedCancel(){
	
	if (CountFlag() == 1) {
		CDialogEx::OnCancel();
	}else {
		MessageBox(TEXT("请至少勾选一个项目吧？"));
	}

}
