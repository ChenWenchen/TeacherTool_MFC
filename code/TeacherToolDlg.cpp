// TeacherToolDlg.cpp: 实现文件//
#include "pch.h"
#include "framework.h"
#include "TeacherTool.h"
#include "TeacherToolDlg.h"
#include "afxdialogex.h"
#include "CDlg_CreateNewGroup.h"
#include "CWorkbook.h"              //管理单个工作表
#include "CWorkbooks.h"        		//统管所有的工作簿
#include "CApplication.h"        	//Excel应用程序类,管理我们打开的这整个Excel应用
#include "CRange.h"                 //区域类，对EXcel的大部分操作都要和这个打招呼
#include "CWorksheet.h"             //工作薄中的单个工作表
#include "CWorksheets.h"            //统管当前工作簿中的所有工作表
#include <fstream>
#include <afxpriv.h>
#include "myExcelReader.h"
using namespace std;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define USE_MODE1 1 // 名字+话术
#define USE_MODE2 2 // 名字+成绩+评价
#define USE_MODE3 3 // 名字+评价

///////////////数据分析相关变量///////////////////
///////////////数据分析相关变量///////////////////
///////////////数据分析相关变量///////////////////

int studentNum;				//学生数量
int qNum;					//原始小分数量（题目数量）

CStringArray* tableHead;    //表头
CStringArray* studentName;	//学生姓名数组，大小是studentNum
CStringArray* studentName2; //去掉姓名的学生名字数组
float* studentSc;			//学生原始成绩数组(一维)，每个学生有qNum个分数，第i个学生的第j个分表示为StudentSc[i*qNum+j]

int speakNum;				//话术数量（话术excel的行数）
CStringArray* speakArr;		//存话术字符串数组

int speakKind;				//如果mode=3，则此变量代表有多少个评价维度

int groupNum;				//有多少个自定义题组模块
CStringArray* groupName;	//存自定义题组模块的名字(大小在初始化函数初始化为100)

int** groupFlagArr;			//存题组模块标志的二维数组，groupFlagArr[i]表示第i个题组，每个groupFlagArr[i]是一个长度为qNum的一维数组
float** groupSc;			//存每个学生每个模块统计出的分数  共有studentNum个int*，每个int数组长度为groupNum

CStringArray* reportArr;	//生成的最终评价话术

int doneCal;				//是否已经计算完分数

int ifFinishedHtml;			//是否已经导出成html 

int ifReplace;              //是否将回车换行替换成占位符

int mode;					//=1表示只用于名字+话术；=2表示名字+成绩+评价；=3表示名字+评价



// CTeacherToolDlg 对话框(程序主对话框)
CTeacherToolDlg::CTeacherToolDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_TEACHERTOOL_DIALOG, pParent), v_ifRemoveFirst(FALSE)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);//载入自定义图标
}

void CTeacherToolDlg::DoDataExchange(CDataExchange* pDX){
	CDialogEx::DoDataExchange(pDX);
	//DDX_Control(pDX, IDC_EDIT1, v_question_num);
	DDX_Control(pDX, IDC_COMBO2, v_groupList);
	DDX_Control(pDX, IDC_ADD_GROUPNAME, v_newGroupName);
	DDX_Control(pDX, IDC_REPORT_TABLE, v_reportTable);
	//DDX_Control(pDX, IDC_REMOVE_FIRST, v_ifRemove);
	//DDX_Control(pDX, IDC_EDIT1, v_question_num);
	//DDX_Control(pDX, IDC_REMOVE_NOT_JUDGE, v_notJudge);
	DDX_Check(pDX, IDC_REMOVE_FIRST, v_ifRemoveFirst);
}

BEGIN_MESSAGE_MAP(CTeacherToolDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_IMPORT_STUDENT, &CTeacherToolDlg::OnBnClickedImportStu)
	ON_BN_CLICKED(IDC_IMPORT_SPEAK, &CTeacherToolDlg::OnBnClickedImportSpk)
	ON_BN_CLICKED(IDC_ADD_GROUP_BNT, &CTeacherToolDlg::OnBnClickedAddGroupBnt)
	//ON_EN_SETFOCUS(IDC_EDIT1, &CTeacherToolDlg::OnEnSetfocusEdit1)
	//ON_EN_KILLFOCUS(IDC_EDIT1, &CTeacherToolDlg::OnEnKillfocusEdit1)
	ON_CBN_SELCHANGE(IDC_COMBO2, &CTeacherToolDlg::OnCbnSelchangeCombo2)
	ON_BN_CLICKED(IDC_COUNT_GROUP_BNT, &CTeacherToolDlg::OnBnClickedCountGroupBnt)
	ON_BN_CLICKED(IDC_GENERATE_REPORT, &CTeacherToolDlg::OnBnClickedGenerateReport)
	//ON_BN_CLICKED(IDC_REMOVE_FIRST, &CTeacherToolDlg::OnBnClickedRemoveFirst)
	//ON_BN_CLICKED(IDC_REMOVE_NOT_JUDGE, &CTeacherToolDlg::OnBnClickedRemoveNotJudge)
	ON_BN_CLICKED(IDC_EXPORT_EXCEL, &CTeacherToolDlg::OnBnClickedExportExcel)
	ON_BN_CLICKED(IDC_EXPORT_TXT, &CTeacherToolDlg::OnBnClickedExportTxt)
	ON_BN_CLICKED(IDC_BUTTON_LINKWEB, &CTeacherToolDlg::OnBnClickedLinkWeb)
END_MESSAGE_MAP()

// CTeacherToolDlg 消息处理程序
BOOL CTeacherToolDlg::OnInitDialog(){
	CDialogEx::OnInitDialog();
	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr){
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty()){
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}
	SetWindowText(L"学生评价生成绩PC版2.0");
	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标
	/////////////////////////////////////////
	//////////////部分数据的初始化///////////
	/////////////////////////////////////////
	SetAddGroupVisble(FALSE);//默认禁用“添加题组模块的按钮”“按模块统计”

	qNum = -1;							//题目数量初始化
	qNum = -1;							//题目数量初始化
	studentNum = -1;					//初始化学生数量用来检错

	groupNum = 0;						//初始化自定义题型组的数量
	groupFlagArr = new int* [100];		// 最多自定义100个题组/模块够了吧
	groupName = new CStringArray[100];

	speakNum = -1;

	doneCal = -1;						//未分组计算成绩
	ifFinishedHtml = 0;					//未导出成html
	mode = USE_MODE2;					//默认使用名字+成绩+评价功能
	speakKind = 0;						//评价维度个数
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CTeacherToolDlg::OnSysCommand(UINT nID, LPARAM lParam){
	if ((nID & 0xFFF0) == IDM_ABOUTBOX){
		//CAboutDlg dlgAbout;//dlgAbout.DoModal();
	}else{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，这将由框架自动完成。
void CTeacherToolDlg::OnPaint(){
	if (IsIconic()){
		CPaintDC dc(this); // 用于绘制的设备上下文
		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);
		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;
		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}else{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标显示。
HCURSOR CTeacherToolDlg::OnQueryDragIcon(){
	return static_cast<HCURSOR>(m_hIcon);
}

///////////////////////////////////////////////////////
/////////// 导入学生成绩数据///////////////////////////
///////////////////////////////////////////////////////
void CTeacherToolDlg::OnBnClickedImportStu(){
	
	tableHead = new CStringArray;  //存表头数组
	studentName = new CStringArray;//存名字数组
	studentName2 = new CStringArray;
	studentSc = new float[100000]; //存分数数组
	UpdateData(TRUE);//更新"去掉姓名"勾选框状态
	
	clearTable();

	//readStudent函数负责导入学生成绩excel，写入表头、学生名字、学生成绩这三个数组
	myExcelReader r = myExcelReader();
	if (!r.readStudent(&studentNum, &qNum, tableHead, studentName,studentName2, studentSc, v_ifRemoveFirst)) {
		MessageBox(TEXT("读入失败~"));
		return;//读入失败则不用更新列表显示
	}

	if (qNum == -1) {
		MessageBox(TEXT("貌似导入失败哦~可能是导错表格或者表格有空值，另外可以把没有数据的行列删除一下（“删除”，不是“清空内容”）"));
		return;
	}else if (qNum == 0) {
		MessageBox(TEXT("数据只有一列，待会只能进行话术合成，不能分析成绩哦~"));
		mode = USE_MODE1;			//表示只能使用名字+话术功能
	}else {
		if (IDYES == ::MessageBox(NULL, L"请问导入的是“成绩分数”还是“评价等级”呢？（点击是代表成绩分数）", L"询问", MB_YESNO)) {
			mode = USE_MODE2;		//代表使用姓名+成绩分析功能
			SetAddGroupVisble(TRUE);//取消禁用“添加题组模块的按钮”“按模块统计”
		}else {
			mode = USE_MODE3;		//代表使用姓名+评价生成功能
		}
	}

	if (mode != USE_MODE2){			//除了mode2下多次读入学生数据可以沿用原来的题组模块，其余情况都将已添加的模块清空
		groupNum = 0;						//初始化自定义题型组的数量
		groupFlagArr = new int* [100];		// 最多自定义100个题组/模块够了吧
	;
		groupName = new CStringArray[100];
		v_groupList.ResetContent();
		SetAddGroupVisble(FALSE);//禁用“添加题组模块的按钮”“按模块统计”
	}

	//渲染表格显示excel内容
	for (int i = 0; i < qNum+1; i++) {							//构造表头
		v_reportTable.InsertColumn(i + 1, tableHead->GetAt(i), LVCFMT_LEFT, 80);
	}
	CString scoreTemp;
	for (int i = 0; i < studentNum; i++) {						//逐行填入学生数据（表格第二行开始）
		v_reportTable.InsertItem(i, studentName->GetAt(i));		//向表格组件中插入一个行
		for (int j = 0; j < qNum; j++) {						//对这个行，逐列填充分数
			scoreTemp.Format(TEXT("%.2f"), studentSc[i * qNum + j]);
			v_reportTable.SetItemText(i, j + 1, scoreTemp);		//给第i行的这个item设置第j+1列数据
		}
	}

	//设置额外样式属性，整行选中状态，显示白色框
	v_reportTable.SetExtendedStyle(v_reportTable.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

	speakNum = -1;//多次上传学生名字后要重新上传话术

	return;
}

///////////////////////////////////////////////////////
/////////// 导入学生话术列表///////////////////////////
///////////////////////////////////////////////////////
void CTeacherToolDlg::OnBnClickedImportSpk(){
	if (studentNum == -1) {
		MessageBox(TEXT("请先导入学生成绩吧~"));
		return;
	}

	speakArr = new CStringArray;//存话术数组
	//readSpeak函数负责导入话术数据，更新话术数量speakNum ， 更新话术维度个数speakKind，话术存入speakArr
	myExcelReader r = myExcelReader();
	if (!r.readSpeak(&speakNum, &speakKind, mode, speakArr)) {
		MessageBox(TEXT("读入话术失败~"));
		speakNum=-1;//读入失败重置参数
		speakKind = 0;
	}

	//////////打开xlsx或xls文件//////////////
	//HRESULT hr;					//HRESULT函数返回值
	//hr = CoInitialize(NULL);	//CoInitialize用来告诉 Windows以单线程的方式创建com对象
	//if (FAILED(hr)) { AfxMessageBox(_T("Failed to call Coinitialize()")); }
	//CFileDialog  filedlg(TRUE, L"*.xl*", NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, L"Xlsx文件 (*.xl*)|*.xl*");
	//filedlg.m_ofn.lpstrTitle = L"打开文件";
	//CString strFilePath;
	//if (IDOK == filedlg.DoModal()) {
	//	strFilePath = filedlg.GetPathName();
	//}
	//else { return; }
	//CApplication app2;	//Excel程序
	//CWorkbooks books2;	//工作簿集合
	//CWorkbook book2;	//工作表
	//CWorksheets sheets2;//工作簿集合
	//CWorksheet sheet2;	//工作表集合
	//CRange range2;		//使用区域
	//CRange iCell2;
	//LPDISPATCH lpDisp2;
	//COleVariant vResult2;  //COleVariant类是对VARIANT结构的封装
	//COleVariant covOptiona2((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	//if (!app2.CreateDispatch(_T("Excel.Application"), NULL)) {
	//	AfxMessageBox(_T("无法启动Excel服务器!"));
	//	return;
	//}
	//books2.AttachDispatch(app2.get_Workbooks());
	//lpDisp2 = books2.Open(strFilePath, covOptiona2, covOptiona2,
	//	covOptiona2, covOptiona2, covOptiona2, covOptiona2, covOptiona2,
	//	covOptiona2, covOptiona2, covOptiona2, covOptiona2, covOptiona2,
	//	covOptiona2, covOptiona2);
	//book2.AttachDispatch(lpDisp2);	//得到Workbook    
	//sheets2.AttachDispatch(book2.get_Worksheets());//得到Worksheets   
	////如果有单元格正处于编辑状态中，此操作不能返回，会一直等待 
	//lpDisp2 = book2.get_ActiveSheet();
	//sheet2.AttachDispatch(lpDisp2);//得到当前活跃sheet 	
	////sheet2.AttachDispatch(sheets2.get_Item(COleVariant((long)2)));//得到第2个sheet

	////获取excel行数列数
	//range2.AttachDispatch(sheet2.get_UsedRange(), TRUE);
	//long rowNum = 0;
	//long colNum = 0;
	//range2.AttachDispatch(range2.get_Rows(), TRUE);
	//rowNum = range2.get_Count();
	//range2.AttachDispatch(range2.get_Columns(), TRUE);
	//colNum = range2.get_Count();
	//CString tmpRowAndCol;
	//tmpRowAndCol.Format(TEXT("读取%ld行%ld列话术数据"), rowNum, colNum);
	//MessageBox(tmpRowAndCol);
	//speakNum = rowNum;
	//////////////////////////////////////////////////////////////////////////////////////////////////////
	//////////////以下是遍历单元格的值并且记录在全局变量//////////////////////////////////////////////////
	//////////////////////////////////////////////////////////////////////////////////////////////////////
	//if (mode == USE_MODE1||mode== USE_MODE2) {
	//	speakArr = new CStringArray[speakNum];//存话术数组
	//	CString resultStr;//用来接收excel单元格数据
	//	for (int i = 1; i <= speakNum; i++) {
	//		range2.AttachDispatch(sheet2.get_Cells());
	//		range2.AttachDispatch(range2.get_Item(COleVariant((long)i), COleVariant((long)2)).pdispVal); //读excel中第i行第j列
	//		vResult2 = range2.get_Value2();
	//		if (vResult2.vt == VT_BSTR) {//该格数据是字符串  
	//			resultStr = vResult2.bstrVal;
	//			speakArr->Add(resultStr);
	//		}else if (vResult2.vt == VT_R8) { //该个数据是8字节的数字  
	//			resultStr.Format(TEXT("%.2lf"), vResult2.dblVal);
	//			speakArr->Add(resultStr);
	//		}
	//	}
	//}else if (mode == USE_MODE3) {
	//	speakKind = colNum - 1;//评价维度数量
	//	int maxRowNum = rowNum - 1;
	//	int dimNum = colNum - 1;
	//	speakArr = new CStringArray[maxRowNum*dimNum];
	//	CString resultStr;
	//	for (int i = 2; i <= rowNum; i++) { //mode3从第二行第二列开始读
	//		for(int j = 2; j <= colNum; j++) {
	//			range2.AttachDispatch(sheet2.get_Cells());
	//			range2.AttachDispatch(range2.get_Item(COleVariant((long)i), COleVariant((long)j)).pdispVal); //读excel中第i行第j列
	//			vResult2 = range2.get_Value2();
	//			if (vResult2.vt == VT_BSTR) {//该格数据是字符串  
	//				resultStr = vResult2.bstrVal;
	//				speakArr->Add(resultStr);
	//			}else if (vResult2.vt == VT_R8) { //该个数据是8字节的数字  
	//				resultStr.Format(TEXT("%.2lf"), vResult2.dblVal);
	//				speakArr->Add(resultStr);
	//			}else if (vResult2.vt == VT_EMPTY) {   //数据为空
	//				CString empty;
	//				empty.Format(TEXT("(%d,%d)处话术为空",i,j));
	//				speakArr->Add(empty);
	//			}
	//		}
	//	}
	//}
	//// 释放对象      
	//range2.ReleaseDispatch();
	//sheet2.ReleaseDispatch();
	//sheets2.ReleaseDispatch();
	//book2.ReleaseDispatch();
	//books2.ReleaseDispatch();
	//books2.Close();
	//app2.Quit();//app必须先quit再release否则excel文件会一直处于编辑状态被锁定
	//app2.ReleaseDispatch();
	//app2.put_Visible(TRUE);
	//app2.put_UserControl(TRUE);

	//多次导入话术后要重新按模块统计和生成html
	doneCal = -1;						//未分组计算成绩，无法按生成网页报告
	ifFinishedHtml = 0;					//未生成html,无法按导出到excel
}

///////////////////////////////////////////////////////
/////////// 添加题组模块///////////////////////////////
///////////////////////////////////////////////////////
void CTeacherToolDlg::OnBnClickedAddGroupBnt(){

	CString groupNameAdd;
	v_newGroupName.GetWindowText(groupNameAdd);
	if(mode==1) {
		MessageBox(TEXT("因为只导入了学生名字，所以不能分析成绩，请确认导入话术列表后直接点击“生成网页报告”吧~"));
	}else if (mode == 3) {
		MessageBox(TEXT("因为导入的是学生评价等级，不是分析成绩，所以无法添加题组~"));
	}else if (studentNum == -1) {
		MessageBox(TEXT("未导入学生成绩哦~"));
	}else if (groupNameAdd == "") {
		MessageBox(TEXT("题型名为空不太好吧"));
	}else if(speakNum == -1){
		MessageBox(TEXT("请先导入话术吧~"));
	}else {
		v_groupList.AddString(groupNameAdd);		//列表中加入新建的题型
		v_newGroupName.SetWindowTextW(TEXT(""));	//清空新题型编辑框
		groupNum++;									//多一个自定义的题组

		int listCnt = v_groupList.GetCount();		//获取下拉框内有多少个题型组
		v_groupList.SetCurSel(listCnt - 1);			//设置默认选中最新
		groupFlagArr[listCnt - 1] = new int[qNum];	//为该题型模块新建一个标志数组
		//把题目数量和名字传参过去新对话框动态生成勾选框，groupFlag传入存这个题组/模块包含哪些题目
		CDlg_CreateNewGroup newgroupdlg(qNum, &tableHead, groupFlagArr[listCnt - 1] , groupNameAdd);
		newgroupdlg.DoModal();

		GetDlgItem(IDC_COUNT_GROUP_BNT)->EnableWindow(TRUE);//取消禁用按模块统计
	}
}

///////////////////////////////////////////////////////
/////////// 点击了“按模块统计”///////////////////////////////
///////////////////////////////////////////////////////
void CTeacherToolDlg::OnBnClickedCountGroupBnt(){

	if (mode == 1) {
		MessageBox(TEXT("因为只导入了学生名字，所以不能分析成绩，请确认导入话术列表后点击“生成网页报告”吧~"));
	}else if (mode == 3) {
		MessageBox(TEXT("因为导入的是学生评价等级，不是分析成绩，所以无法添加题组~"));
	}else if (studentNum == -1) {
		MessageBox(TEXT("未导入学生成绩哦~"));
	}else if (mode == 2 && speakNum == -1) {
		MessageBox(TEXT("请导入话术吧~"));
	}else if (groupNum == 0) {
		MessageBox(TEXT("未添加任何模块哦~"));
	}else{
		doCalculate();
		doneCal = 1;
	}
}

///////////////////////////////////////////////////////
/////////// 生成网页报告///////////////////////////////
///////////////////////////////////////////////////////
void CTeacherToolDlg::OnBnClickedGenerateReport(){
	if (mode==2 && doneCal != 1) { //mode =2 才需要先统计模块分
		MessageBox(TEXT("请先导入学生成绩，并且按模块统计分数！"));
		return;
	}else if (speakNum == -1) {
		MessageBox(TEXT("请导入话术！"));
		return;
	}
		
	ofstream file("评价报告.html");

	CString preHtml("<head><link href='./res/index.css' rel='stylesheet' type='text/css'/></head>");
	USES_CONVERSION;
	LPSTR bufff = T2A(preHtml);
	file.write(bufff, strlen(bufff));

	reportArr = new CStringArray[studentNum];
	for (int i = 0; i < studentNum; i++) {
		
		CString easyHead;//easyHead是用来快速复制名字的内容
		easyHead.Format(TEXT("<div class='studentBox' id='s%d'><button class='bntKill'>清除</button><textarea style='overflow:scroll; width:100px; height:150px' id='Name%d' type='text’ class=‘outcome1'>%s</textarea><button class='btn b1' data-clipboard-action='copy' data-clipboard-target='#Name%d'> 复制</button>"),i, i, studentName->GetAt(i), i);
		
		CString head;
		head.Format(TEXT("<textarea style='overflow:scroll; width:350px; height:150px' id='Report%d' type='text’ class=‘outcome2'>"), i);
		
		CString msg;
		CString helloWord;////////需要自定义
		if (mode != 3) {//mode3不需要这个开头话术，直接设置一个全1评价即可
			helloWord = speakArr->GetAt(0);
		}

		msg = msg + helloWord;//加问好（通用语）

		if ( mode == 2) {//循环加模块名字和成绩
			CString thisGName;
			CString thisScore;
			CString general1("方面的得分：");
			for (int g = 0; g < groupNum; g++) {
				v_groupList.GetLBText(g, thisGName);//获取模块名字
				float rate = (groupSc[i][g] / groupSc[studentNum - 1][g]) * 100;//计算得分率
				thisScore.Format(TEXT("%.1f%%."), rate);
				msg = msg + thisGName + general1 + thisScore;//拼接 “xxx方面的得分：98.2%”
				CString comment;
				if (groupSc[i][g] == groupSc[studentNum - 1][g]) {//满分
					comment = speakArr->GetAt(1); 
				}else if (rate >= 90) {
					comment = speakArr->GetAt(2); 
				}else if (rate >= 80) {
					comment = speakArr->GetAt(3);
				}else if (rate >= 70) {
					comment = speakArr->GetAt(4);
				}else if (rate >= 60) {
					comment = speakArr->GetAt(5);
				}else if (rate >= 50) {
					comment = speakArr->GetAt(6);
				}else if (rate >= 40){
					comment = speakArr->GetAt(7);
				}else if (rate >= 30){
					comment = speakArr->GetAt(8);
				}else if (rate >= 20) {
					comment = speakArr->GetAt(9);
				}else if (rate >= 10) {
					comment = speakArr->GetAt(10);
				}else{
					comment = speakArr->GetAt(11);
				}
					msg = msg + comment;
			}
		}else if(mode==3){
			CString thisSpeak;
			for (int d = 0; d < speakKind; d++) {
				int kindSc = studentSc[i * qNum + d];//该学生在这个评价维度下的分数，对应Sc数组里的第i行第d列的分数
				if (kindSc != 0) {//如果分数为0表示这个学生的这一项没有评价
					thisSpeak = speakArr->GetAt((kindSc - 1) * speakKind + d);//用评价等级作为索引找出对应话术
					msg += thisSpeak;
				}
			}
		}

		CString general2;
		if ((mode == 2||mode==1) && speakNum > 0){//如果有导入话术则末尾加入“通用术语（结尾）”
			general2 = speakArr->GetAt(speakArr->GetCount() - 1);
			msg = msg + general2;
		}

		msg.Replace(L"{空格}", L"&nbsp");
		msg.Replace(L"{换行}", L"&#10");
		msg.Replace(L"{name}", studentName->GetAt(i));
		msg.Replace(L"{name2}", studentName2->GetAt(i));

		reportArr->Add(msg);
		
		CString tail;
		tail.Format(TEXT("</textarea><button class='btn b2' data-clipboard-action='copy' data-clipboard-target='#Report%d'> 复制</button></div>"), i);

		CString totalHtml;
		if(msg != ""){ totalHtml = easyHead + head + msg + tail; }
		//USES_CONVERSION;
		LPSTR buf = T2A(totalHtml);
		file.write(buf, strlen(buf));
	}
	
	CString sEnd;
	sEnd.Format(TEXT("<script src='./res/clipboard.min.js'></script><script src='./res/index.js'></script><script>var clipboard = new ClipboardJS('.btn'); </script>"));
	//USES_CONVERSION;
	LPSTR bufEnd = T2A(sEnd);
	file.write(bufEnd, strlen(bufEnd));
	file.close();

	ifFinishedHtml = 1;
	MessageBox(TEXT("报告已经生成到程序所在目录下的“评价报告.html”，用浏览器打开即可~（请及时改名，否则下次会重新刷写）另外请注意一键复制的功能要确保html文件在同目录下有res文件夹"));
	system("start explorer .\\评价报告.html");

	//showReport();//将生成的话术显示在reportTable中
}

///////////////////////////////////////////////////////
/////////// 导出excel报告（用wxPython）////////////////
///////////////////////////////////////////////////////
void CTeacherToolDlg::OnBnClickedExportExcel() {

	if (ifFinishedHtml != 1) {
		MessageBox(TEXT("请先生成网页报告！"));
		return;
	}

	////////打开xlsx或xls文件//////////////
	HRESULT hr;					//HRESULT函数返回值
	hr = CoInitialize(NULL);	//CoInitialize用来告诉 Windows以单线程的方式创建com对象
	if (FAILED(hr)) { AfxMessageBox(_T("Failed to call Coinitialize()")); }
	CApplication app3;	//Excel程序
	CWorkbooks books3;	//工作簿集合
	CWorkbook book3;	//工作表
	CWorksheets sheets3;//工作簿集合
	CWorksheet sheet3;	//工作表集合
	CRange range3;		//使用区域
	CRange iCell3;
	COleVariant vResult3;  //COleVariant类是对VARIANT结构的封装
	COleVariant covOptiona3((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app3.CreateDispatch(_T("Excel.Application"), NULL)) {
		AfxMessageBox(_T("无法创建Excel应用!"));
		return;
	}
	books3 = app3.get_Workbooks();
	book3 = books3.Add(covOptiona3);
	sheets3 = book3.get_Worksheets();
	sheet3 = sheets3.get_Item(COleVariant((short)1));

	CString A;
	CString B;
	CString r;
	for (int i = 0; i < studentNum - 1; i++) {
		r = reportArr->GetAt(i);

		if (r == "")continue;//空评价不导出

		r.Replace(TEXT("&nbsp"), TEXT("{b}"));
		r.Replace(TEXT("&#10"), TEXT("{n}"));
		
		A.Format(TEXT("A%d"), i + 1);
		range3 = sheet3.get_Range(COleVariant(A), COleVariant(A));
		range3.put_Value2(COleVariant(studentName->GetAt(i)));
		
		B.Format(TEXT("B%d"), i + 1);
		range3 = sheet3.get_Range(COleVariant(B), COleVariant(B));
		range3.put_Value2(COleVariant(r));
	}
	//range2.put_Formula(COleVariant(L"=RAND()*1000"));
	CRange cols;
	cols = range3.get_EntireColumn();
	cols.AutoFit();

	// 释放对象      
	cols.ReleaseDispatch();
	range3.ReleaseDispatch();
	sheet3.ReleaseDispatch();
	sheets3.ReleaseDispatch();
	book3.ReleaseDispatch();
	books3.ReleaseDispatch();
	books3.Close();
	app3.Quit();//app必须先quit再release否则excel文件会一直处于编辑状态被锁定
	app3.ReleaseDispatch();
	app3.put_Visible(TRUE);
	app3.put_UserControl(TRUE);

	ShowWindow(SW_MINIMIZE);//最小化程序窗口否则会挡住保存提示
}

///////////////////////////////////////////////////////
/////////// 导出txt文档（用PC微信robot）（现已弃用）////////////////
///////////////////////////////////////////////////////
void CTeacherToolDlg::OnBnClickedExportTxt() {
	if (ifFinishedHtml != 1) {
		MessageBox(TEXT("请先生成网页报告！"));
		return;
	}
	CStdioFile file;
	char* old_locale = _strdup(setlocale(LC_CTYPE, NULL));
	setlocale(LC_CTYPE, "chs");
	file.Open(L"发送清单.txt", CStdioFile::modeCreate | CStdioFile::modeWrite);

	CString tmpRep;
	for (int i = 0; i < studentNum - 1; i++) {
		file.WriteString(studentName->GetAt(i) + "\n");

		tmpRep = reportArr->GetAt(i);
		tmpRep.Replace(TEXT("&nbsp"), TEXT(""));
		tmpRep.Replace(TEXT("&#10"), TEXT(""));
		file.WriteString(tmpRep + "\n");
	}
	file.Close();
	setlocale(LC_CTYPE, old_locale); //还原语言区域的设置 
	free(old_locale);//还原区域设定	
	MessageBox(TEXT("已经生成在该目录下的“发送清单.txt”，用pc版微信发送助手打开即可~"));
}

//按模块统计
void CTeacherToolDlg::doCalculate(){
	//初始化记录学生模块成绩的二维数组
	groupSc = new float* [studentNum];
	for (int i = 0; i < studentNum; i++) {//逐个学生创建一个长度为groupNum的一维数组
		groupSc[i] = new float[groupNum];
	}
	for (int g = 0; g < groupNum; g++) {//逐个自定义题组
		CString strtmp2;
		v_groupList.GetLBText(g, strtmp2);//获取自定义题型的名字
		v_reportTable.InsertColumn(1, strtmp2, LVCFMT_LEFT, 80);//在姓名旁边插入新列
		for (int s = 0; s < studentNum; s++) {//逐个学生
			float score = 0;
			CString sctmp;
			for (int q = 0; q < qNum; q++) {//按groupFlag统计这个学生在该提醒组下的分数
				score = score + studentSc[s * qNum + q] * groupFlagArr[g][q];
				sctmp.Format(TEXT("%.2f"), score);
				groupSc[s][g] = score;
				v_reportTable.SetItemText(s, 1, sctmp);//给第i行的这个item设置第j+1列数据
			}
		}
	}
}

//题组下拉框中选择了新的题组，用于编辑已有题组
void CTeacherToolDlg::OnCbnSelchangeCombo2() {

	int index = v_groupList.GetCurSel(); //获取当前选取的索引位置
	CString name;						
	v_groupList.GetLBText(index, name);  //获取当前选取的题型的名字
	CString msg;
	msg.Format(TEXT("是否要修改\"%s\"这个模块？"), name);

	if (IDYES == ::MessageBox(NULL, msg, L"询问", MB_YESNO)) {
		//把题目数量和名字传参过去新对话框动态生成勾选框，groupFlag传入存这个题组/模块包含哪些题目
		if (qNum > 0) {
			CDlg_CreateNewGroup newgroupdlg(qNum, &tableHead, groupFlagArr[index], name);
			newgroupdlg.DoModal();
		}else {
			MessageBox(L"学生数据未导入成功~暂时不能修改,请导入学生成绩");
		}
	}else {
		MessageBox(L"那你乱按啥");
	}
}

//显示合成好的最终话术
void CTeacherToolDlg::showReport(){

	clearTable();

	//构造表头
	v_reportTable.InsertColumn(1, TEXT("最终话术"), LVCFMT_LEFT, 800);

	//渲染表格显示excel内容
	CString scoreTemp;
	for (int i = 0; i < studentNum; i++) {						//逐行填入学生数据（表格第二行开始）
		v_reportTable.InsertItem(i, reportArr->GetAt(i));		//向表格组件中插入一个行
		//for (int j = 0; j < qNum; j++) {						//对这个行，逐列填充分数
		//	scoreTemp.Format(TEXT("%.2f"), studentSc[i * qNum + j]);
		//	v_reportTable.SetItemText(i, j + 1, scoreTemp);		//给第i行的这个item设置第j+1列数据
		//}
	}

	//设置额外样式属性，整行选中状态，显示白色框
	v_reportTable.SetExtendedStyle(v_reportTable.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

}

//链接到我的花园
void CTeacherToolDlg::OnBnClickedLinkWeb() {
	system("start explorer http://app.chenyuxian.vip:8081/apk/index.html");
}

//清空已显示的列表
void CTeacherToolDlg::clearTable() {
	int cnt = v_reportTable.GetHeaderCtrl()->GetItemCount();
	for (int i = 0; i < cnt; i++) {
		v_reportTable.DeleteColumn(0);
	}
	v_reportTable.DeleteAllItems();//删除干净，不加这句会有残留
}

//设置“添加题组模块”“按模块统计”是否可用
void CTeacherToolDlg::SetAddGroupVisble(BOOL ifVisble) {
	GetDlgItem(IDC_ADD_GROUP_BNT)->EnableWindow(ifVisble);//默认禁用添加题组模块，只有mode=2才打开
	GetDlgItem(IDC_COUNT_GROUP_BNT)->EnableWindow(ifVisble);//禁用用按模块统计，只有mode=2才打开
	GetDlgItem(IDC_ADD_GROUP_BNT)->ShowWindow(ifVisble);
	GetDlgItem(IDC_COUNT_GROUP_BNT)->ShowWindow(ifVisble);
	GetDlgItem(IDC_ADD_GROUPNAME)->ShowWindow(ifVisble);
	GetDlgItem(IDC_COMBO2)->ShowWindow(ifVisble);
	GetDlgItem(IDC_STATIC_STEP3)->ShowWindow(ifVisble);
}

//注释掉回车会关闭窗口的功能
void CTeacherToolDlg::OnOK(){
	//CDialogEx::OnOK();
}

