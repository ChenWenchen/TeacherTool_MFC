#pragma once

class myExcelReader{

	public:
		BOOL readStudent(int* pStudentNum,
					 int* pqNum, 
					 CStringArray* tableHead, 
					 CStringArray* studentName, 
				     CStringArray* studentName2,
					 float* studentSc,
				 	 bool ifRemoveFirst
		);

		BOOL readSpeak(int* speakNum, int* speakKind, int mode, CStringArray* speakArr);

		void reportErro(int* pStudentNum, int* pqNum, int* pBreakFlag, int i, int j);
};

