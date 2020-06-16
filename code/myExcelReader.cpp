#include "pch.h"
#include "framework.h"
#include "myExcelReader.h"
#include "afxdialogex.h"
#include "CWorkbook.h"              //管理单个工作表
#include "CWorkbooks.h"        		//统管所有的工作簿
#include "CApplication.h"        	//Excel应用程序类,管理我们打开的这整个Excel应用
#include "CRange.h"            	    //区域类，对EXcel的大部分操作都要和这个打招呼
#include "CWorksheet.h"             //工作薄中的单个工作表
#include "CWorksheets.h"            //统管当前工作簿中的所有工作表


////读入学生成绩专用函数
BOOL myExcelReader::readStudent(int* pStudentNum,
								int* pqNum,
								CStringArray* tableHead,
								CStringArray* studentName, 
								CStringArray* studentName2,
								float* studentSc,
								bool ifRemoveFirst
	){
	//////////////打开xlsx或xls文件//////////////
	HRESULT hr = CoInitialize(NULL);//HRESULT函数返回值//CoInitialize用来告诉 Windows以单线程的方式创建com对象
	if (FAILED(hr)){ AfxMessageBox(_T("Failed to call Coinitialize()")); }
	CFileDialog  filedlg(TRUE, L"*.xl*", NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, L"Xlsx文件 (*.xl*)|*.xl*");
	filedlg.m_ofn.lpstrTitle = L"打开文件";
	CString strFilePath;
	if (IDOK == filedlg.DoModal()) {
		strFilePath = filedlg.GetPathName();
	}else{ 
		*pStudentNum = -1; //重置这两个参数，相当于未导入成绩
		*pqNum = -1;
		return FALSE; 
	}
	CApplication app1;	//Excel程序
	CWorkbooks books;	//工作簿集合
	CWorkbook book;		//工作表
	CWorksheets sheets; //工作簿集合
	CWorksheet sheet;	//工作表集合
	CRange range;		//使用区域
	CRange iCell;
	LPDISPATCH lpDisp;
	COleVariant vResult;//COleVariant类是对VARIANT结构的封装
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app1.CreateDispatch(_T("Excel.Application"), NULL)) {
		AfxMessageBox(_T("无法启动Excel服务器!"));
		*pStudentNum = -1; //重置这两个参数，相当于未导入成绩
		*pqNum = -1;
		return FALSE;
	}
	books.AttachDispatch(app1.get_Workbooks());
	lpDisp = books.Open(strFilePath, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional, covOptional, covOptional, covOptional,
		covOptional, covOptional);
	//得到Workbook    
	book.AttachDispatch(lpDisp);
	//得到Worksheets   
	sheets.AttachDispatch(book.get_Worksheets());
	//sheet = sheets.get_Item(COleVariant((short)1));
	//得到当前活跃sheet ，如果有单元格正处于编辑状态中，此操作不能返回，会一直等待 
	lpDisp = book.get_ActiveSheet();
	sheet.AttachDispatch(lpDisp);
	//获取excel行数列数
	range.AttachDispatch(sheet.get_UsedRange(), TRUE);
	range.AttachDispatch(range.get_Rows(), TRUE);
	int rowNum = range.get_Count();
	range.AttachDispatch(range.get_Columns(), TRUE);
	int  colNum = range.get_Count();

	///////////////////////////////读取行列数////////////////////////////////////
	CString tmpRowAndCol;
	tmpRowAndCol.Format(TEXT("读取%ld行%ld列学生数据,请稍等~"), rowNum, colNum);
	AfxMessageBox(tmpRowAndCol);
	
	//行数-1（去掉表头）为学生人数，列数减-1（去掉名字）是项目数
	*pStudentNum = rowNum - 1;
	*pqNum = colNum - 1;

	if (*pStudentNum < 1) {
		AfxMessageBox(TEXT("表格好像是空的吧？请检查excel"));
		*pStudentNum = -1;
		*pqNum = -1;
	}else {
		////////////以下是遍历单元格的值并且记录在主dlg的全局变量////////////////
		int studentNum = *pStudentNum;
		int qNum = *pqNum;
		int k = 0;									//存分索引
		int breakFlag = 0;							//读到空值时跳出双循环结束读取
		CString resultStr;							//用来接收excel单元格数据
		CString firstC;
		for (int i = 1; i <= rowNum; i++) {			//遍历excel的所有行
			for (int j = 1; j <= colNum; j++) {     //遍历所有列
				range.AttachDispatch(sheet.get_Cells());
				range.AttachDispatch(range.get_Item(COleVariant((long)i), COleVariant((long)j)).pdispVal); //读excel中第i行第j列
				vResult = range.get_Value2();

				if (i == 1) { //第一行是表头
					if (vResult.vt == VT_BSTR) {	    //该格数据是字符串  
						tableHead->Add(vResult.bstrVal);
					}else if (vResult.vt == VT_R8) {    //表头是纯数字
						CString numHead;
						numHead.Format(TEXT("%.2f"), vResult.dblVal);
						tableHead->Add(numHead);
					}else if (vResult.vt == VT_EMPTY) {   //表头数据为空
						reportErro(pStudentNum, pqNum, &breakFlag, i, j);
						break;
					}
				}else if (j == 1) {	//第一列是姓名（第一行第一列）
					if (vResult.vt == VT_BSTR) {		 //该格数据是字符串  
						resultStr = vResult.bstrVal;
						if (resultStr == "") {			 //网页版我的花园可能会出现类型为bstrVal但是空值
							reportErro(pStudentNum, pqNum, &breakFlag, i, j);
							break;
						}
						char c = (resultStr.GetAt(resultStr.GetLength() - 1));	//获取最后一个字符
						if (c >= '0' && c <= '9') {			//末尾是数字则去掉，有编码问题
							resultStr = resultStr.Left(resultStr.GetLength() - 1);
						}
						studentName->Add(resultStr);
						//去掉姓氏的名字存在studentName2
						resultStr = resultStr.Right(2);	//只要最右两个名字字符（不兼容欧阳修、慕容复）
						studentName2->Add(resultStr);
					}else {   //名字这一列只能是字符串，否则都报错
						reportErro(pStudentNum, pqNum, &breakFlag, i, j);
						break;
					}
				}else { // 第2行第2列开始应该是中间数据
					if (vResult.vt == VT_BSTR) {	    //我的花园app网页版会出现str类型的数值  
						resultStr = vResult.bstrVal;
						studentSc[k] = _ttof(resultStr);//str转成float存
						k++; //存分索引自加
					}else if (vResult.vt == VT_R8) {    //表头是纯数字
						studentSc[k] = vResult.dblVal;
						k++;//存分索引自加
					}else if (vResult.vt == VT_EMPTY) {   //表头数据为空
						reportErro(pStudentNum, pqNum, &breakFlag, i, j);
						break;
					}
				}
				//	if (vResult.vt == VT_BSTR) {	    //该格数据是字符串  
				//		resultStr = vResult.bstrVal;
				//		AfxMessageBox(resultStr);
				//		if (i == 1) {					//第一行是表头，直接存
				//			tableHead->Add(resultStr);
				//		}else{
				//			if (resultStr == "") {		//网页版我的花园可能会出现类型为bstrVal但是空值
				//				reportErro(pStudentNum, pqNum, &breakFlag, i, j);
				//				break;
				//			}
				//			char c = (resultStr.GetAt(resultStr.GetLength() - 1));	//获取最后一个字符
				//			if (c >= '0' && c <= '9') {			//末尾是数字则去掉，有编码问题
				//				resultStr = resultStr.Left(resultStr.GetLength() - 1);
				//			}
				//			if (ifRemoveFirst) {
				//				resultStr = resultStr.Right(2);	//只要最右两个名字字符（不兼容欧阳修、慕容复）
				//			}
				//			studentName->Add(resultStr);
				//		}
				//	}else if (vResult.vt == VT_R8) {	//该个数据是8字节的数字  
				//		if (i == 1) {					//如果表头是数字要转换成CString存
				//			CString numHead;
				//			numHead.Format(TEXT("%.2f"), vResult.dblVal);
				//			tableHead->Add(numHead);
				//		}else {
				//			studentSc[k] = vResult.dblVal;
				//			k++; //k只用来存成绩分数所以只在这里加加
				//		}
				//	}else if (vResult.vt == VT_EMPTY) {   //数据为空
				//		reportErro(pStudentNum,  pqNum,  &breakFlag, i, j);
				//		break;
				//	}
			}
			if (breakFlag) break;
		}
	}

	//释放对象
	range.ReleaseDispatch();
	sheet.ReleaseDispatch();
	sheets.ReleaseDispatch();
	book.ReleaseDispatch();
	books.ReleaseDispatch();
	books.Close();
	app1.Quit();	//app必须先quit再release否则excel文件会一直处于编辑状态被锁定
	app1.ReleaseDispatch();
	app1.put_Visible(TRUE);
	app1.put_UserControl(TRUE);

	return TRUE;
}



////读入话术专用函数
BOOL myExcelReader::readSpeak(int* speakNum , 
							  int* speakKind, 
							  int mode , 
							  CStringArray* speakArr
	){
	////////打开xlsx或xls文件//////////////
	HRESULT hr;					//HRESULT函数返回值
	hr = CoInitialize(NULL);	//CoInitialize用来告诉 Windows以单线程的方式创建com对象
	if (FAILED(hr)) { AfxMessageBox(_T("Failed to call Coinitialize()")); }
	CFileDialog  filedlg(TRUE, L"*.xl*", NULL, OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT, L"Xlsx文件 (*.xl*)|*.xl*");
	filedlg.m_ofn.lpstrTitle = L"打开文件";
	CString strFilePath;
	if (IDOK == filedlg.DoModal()) {
		strFilePath = filedlg.GetPathName();
	}else{ 
		return FALSE; }
	CApplication app2;	//Excel程序
	CWorkbooks books2;	//工作簿集合
	CWorkbook book2;	//工作表
	CWorksheets sheets2;//工作簿集合
	CWorksheet sheet2;	//工作表集合
	CRange range2;		//使用区域
	CRange iCell2;
	LPDISPATCH lpDisp2;
	COleVariant vResult2;  //COleVariant类是对VARIANT结构的封装
	COleVariant covOptiona2((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	if (!app2.CreateDispatch(_T("Excel.Application"), NULL)) {
		AfxMessageBox(_T("无法启动Excel服务器!"));
		return FALSE;
	}
	books2.AttachDispatch(app2.get_Workbooks());
	lpDisp2 = books2.Open(strFilePath, covOptiona2, covOptiona2,
		covOptiona2, covOptiona2, covOptiona2, covOptiona2, covOptiona2,
		covOptiona2, covOptiona2, covOptiona2, covOptiona2, covOptiona2,
		covOptiona2, covOptiona2);
	book2.AttachDispatch(lpDisp2);	//得到Workbook    
	sheets2.AttachDispatch(book2.get_Worksheets());//得到Worksheets   
	//如果有单元格正处于编辑状态中，此操作不能返回，会一直等待 
	lpDisp2 = book2.get_ActiveSheet();
	sheet2.AttachDispatch(lpDisp2);//得到当前活跃sheet 	
	//sheet2.AttachDispatch(sheets2.get_Item(COleVariant((long)2)));//得到第2个sheet

	//获取excel行数列数
	range2.AttachDispatch(sheet2.get_UsedRange(), TRUE);
	long rowNum = 0;
	long colNum = 0;
	range2.AttachDispatch(range2.get_Rows(), TRUE);
	rowNum = range2.get_Count();
	range2.AttachDispatch(range2.get_Columns(), TRUE);
	colNum = range2.get_Count();
	CString tmpRowAndCol;
	tmpRowAndCol.Format(TEXT("读取%ld行%ld列话术数据"), rowNum, colNum);
	AfxMessageBox(tmpRowAndCol);
	*speakNum = rowNum;
	////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////以下是遍历单元格的值并且记录在全局变量//////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////
	if (mode == 1 || mode == 2) {
		CString resultStr;//用来接收excel单元格数据
		for (int i = 1; i <= *speakNum; i++) {
			range2.AttachDispatch(sheet2.get_Cells());
			range2.AttachDispatch(range2.get_Item(COleVariant((long)i), COleVariant((long)2)).pdispVal); //读excel中第i行第j列
			vResult2 = range2.get_Value2();
			if (vResult2.vt == VT_BSTR) {//该格数据是字符串  
				resultStr = vResult2.bstrVal;
				speakArr->Add(resultStr);
			}
			else if (vResult2.vt == VT_R8) { //该个数据是8字节的数字  
				resultStr.Format(TEXT("%.2lf"), vResult2.dblVal);
				speakArr->Add(resultStr);
			}
		}
	}
	else if (mode == 3) {
		*speakKind = colNum - 1;//评价维度数量
		int maxRowNum = rowNum - 1;
		int dimNum = colNum - 1;
		CString resultStr;
		for (int i = 2; i <= rowNum; i++) { //mode3从第二行第二列开始读
			for (int j = 2; j <= colNum; j++) {
				range2.AttachDispatch(sheet2.get_Cells());
				range2.AttachDispatch(range2.get_Item(COleVariant((long)i), COleVariant((long)j)).pdispVal); //读excel中第i行第j列
				vResult2 = range2.get_Value2();
				if (vResult2.vt == VT_BSTR) {//该格数据是字符串  
					resultStr = vResult2.bstrVal;
					speakArr->Add(resultStr);
				}
				else if (vResult2.vt == VT_R8) { //该个数据是8字节的数字  
					resultStr.Format(TEXT("%.2lf"), vResult2.dblVal);
					speakArr->Add(resultStr);
				}
				else if (vResult2.vt == VT_EMPTY) {   //数据为空
					CString empty(" ");
					//empty.Format(TEXT("(%d,%d)处话术为空", i, j));
					speakArr->Add(empty);
				}
			}
		}
	}
	// 释放对象      
	range2.ReleaseDispatch();
	sheet2.ReleaseDispatch();
	sheets2.ReleaseDispatch();
	book2.ReleaseDispatch();
	books2.ReleaseDispatch();
	books2.Close();
	app2.Quit();//app必须先quit再release否则excel文件会一直处于编辑状态被锁定
	app2.ReleaseDispatch();
	app2.put_Visible(TRUE);
	app2.put_UserControl(TRUE);

	return TRUE;
}



void myExcelReader::reportErro(int* pStudentNum, int* pqNum, int* pBreakFlag , int i, int j){
	CString deb;
	deb.Format(TEXT("表格中有数据出现错误，读入失败! 请检查单元格（%d,%d）"), i, j);
	AfxMessageBox(deb);
	*pStudentNum = -1; //重置这两个参数，相当于未导入成绩
	*pqNum = -1;
	*pBreakFlag = 1;
}



