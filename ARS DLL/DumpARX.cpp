// DumpARX.cpp: implementation of the DumpARX class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "DumpARX.h"
#include "Globals.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#include "RecordList.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CDumpARX::CDumpARX()
{
	iReturnCode = rc_none;
}

CDumpARX::~CDumpARX()
{
}

/////////////////////////////////////////////////////////////////////////////////
// Description: Dumps records for the forms in "formList" to ARX files.
//				If m_tModTime = 0, a complete backup, else inremental
// Returns:		TRUE on success
//				FALSE on failure
BOOL CDumpARX::DumpARXRecords(CARSConnection &arsConnect, CFormList &formList, CString strBackupDir, 
							  CString &strQualifier, 
							  CTime m_tModTime, // zero, if complete backup non-zero if incremental backup
							  UINT uiMaxNum)
{
	POSITION formPOS; // index for formList
	bool bResult; // working result for recodList return value
	CString strModTime; // working string for mod time

//	SECURITY_ATTRIBUTES z;
//	z.nLength = sizeof(z);
//	z.lpSecurityDescriptor = NULL;
//	z.bInheritHandle = FALSE;

//	// Ensure the backup directory exists, if not create it
//	strBackupDir += "\\";
	if(!CreateDirectorySeq(LPCTSTR(strBackupDir) ) ){
		iReturnCode = rc_no_file;
		char szErrorText[255];
		DWORD iError = GetLastError();
		FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, iError, NULL, szErrorText, sizeof(szErrorText), NULL);
		log.Log(CFileLogging::level_lower, "CDumpARS::DumpARSRecords():  Error creating backup directory - ", szErrorText, NULL);
		return FALSE;
	}

	

	// loop once for each form, dumping each form to Backup dir
	for(formPOS = formList.GetHeadPosition(); formPOS != NULL;)
	{	
		// working string, fixed form name
		CString strFixedFormName(ReplaceFormChars( formList.GetAt(formPOS) )); 
		
		// working string to hold backup file name for current form.
		CString strBackupFileName(strBackupDir); 

		// Build the finished Backup Filename
		strBackupFileName += "\\" + strFixedFormName + ".arx";
		CleanupFileName(strBackupFileName);

		// Do setup for incremental backup job
		if(m_tModTime.GetTime() > 0) {
			// add last mod time to qualifier on incremental jobs
			strModTime.Format("('6' >= \"%d/%d/%d %d:%d:%d\") AND", m_tModTime.GetMonth(), m_tModTime.GetDay(),
					m_tModTime.GetYear(), m_tModTime.GetHour(), m_tModTime.GetMinute(), m_tModTime.GetSecond());
		
			strQualifier.Insert(0, "("); // insert ( at beginning of string
			strQualifier.Insert(strQualifier.GetLength() + 10, ")"); // append ) at end of string
			strQualifier.Insert(0, (LPCTSTR)strModTime); // Insert modTime at begin of qualification string
			// end result should be ex: ('6' >= "1/18/1973 "15:30:00") AND ('Status' != "Closed")
			// if strQualifier before is: 'Status' != "Closed"
		}// end if		

		// First, fill the record list with all entries from the form
		CString dumpForm; dumpForm = formList.GetNext(formPOS);
		recordList.FillRecordList(arsConnect, dumpForm, strQualifier, uiMaxNum);

		// don't dump anything if there are no records.
		if(recordList.GetCount() <= 0) {
			continue;
		}

		CStdioFile File;
// ADD open in append or complete backup mode

		// if in complete backup, overwrite the file
		if(m_tModTime.GetTime() == 0) {
			// COMPLETE
			// For each form, create and open a new ARX output file
			if(!File.Open(LPCSTR(strBackupFileName), 
				CFile::modeCreate | CFile::shareDenyWrite | CFile::modeWrite | CFile::typeText) ) {
				// file coulnd't be opened, exit
				iReturnCode = rc_no_file;
				return FALSE;
			}
		}else {
			// INCREMENTAL
			// dump in incremental mode, append, no overwrite to file
			if(!File.Open(LPCSTR(strBackupFileName), 
					CFile::modeCreate | CFile::modeNoTruncate | CFile::shareDenyWrite | \
					CFile::modeWrite | CFile::typeText) ) {
				// file coulnd't be opened, exit
				iReturnCode = rc_no_file;
				return FALSE;
			}
		}
		
// ADD  append / overwrite to recordList.DumpARS for incremental jobs
		// dump the list of records
		bResult = recordList.DumpARX(arsConnect, File, strFixedFormName, m_tModTime);
		// Close the ARX file for this form
		File.Close();

		if(bResult == FALSE)
			return FALSE;
	} // end for

	return TRUE;
}

///////////////////////////////////////////////////////////
// This function will return a cleaned up form name
// with all invalid characters replaced with underscores "_"
CString CDumpARX::ReplaceFormChars(CString &strFormName)
{
	CString strGoodFormName(strFormName); // working form name

	for(int i=0; i<strGoodFormName.GetLength(); i++)
	{
		if((strGoodFormName[i] >= '0' && strGoodFormName[i] <= '9') /*numbers*/
			|| (strGoodFormName[i] >= 'A' && strGoodFormName[i] <= 'Z') /*cap letters*/
			|| (strGoodFormName[i] >= 'a' && strGoodFormName[i] <= 'z') /*lower letters*/)
		{
			continue; // it's a valid character, continue to next character
		}else
		{
			// it's a bad character, convert it to underscore
			strGoodFormName.SetAt(i, '_');
		} // end if/else
	}// end for
	return strGoodFormName;
}

//DEL int CDumpARX::CreateARXFile(CString strFileName)
//DEL {
//DEL 	// Create the file
//DEL 	HANDLE hFile = CreateFile(LPCSTR(strFileName),
//DEL          GENERIC_WRITE, FILE_SHARE_READ,
//DEL          NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
//DEL 	
//DEL 	// check for errors
//DEL 	if (hFile == INVALID_HANDLE_VALUE)
//DEL 	{
//DEL       // Log the error to a log file
//DEL 	  //AfxMessageBox(_T("Couldn't create the file!"));
//DEL 	}
//DEL 
//DEL }

CString CDumpARX::CleanupFileName(CString &strFilename)
{
	while(strFilename.Replace("\\\\", "\\"))
	{
	}

	return strFilename;
}

//DEL void CDumpARX::Stop()
//DEL {
//DEL 	recordList.SetStop();
//DEL }

void CDumpARX::SetStop(bool newbStop)
{
	iReturnCode = rc_stopped; // save job stopped return code
	recordList.SetStop(newbStop);
}

int CDumpARX::GetReturnCode()
{
	return iReturnCode;
}

bool CDumpARX::CreateDirectorySeq(LPCTSTR lpszDirPath)
{
	TCHAR szShorterPath[MAX_PATH];	
	TCHAR szDirectory[MAX_PATH];	
	
	if(!CreateDirectory(lpszDirPath, NULL)){	
		DWORD iError = GetLastError();
#ifdef _DEBUG
		char szErrorText[255];
		FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, iError, NULL, szErrorText, sizeof(szErrorText), NULL);
#endif
		if(iError == 183) // this is the code for the directory already existing.
			return true;
		_tcscpy(szDirectory, lpszDirPath);		
		if(szDirectory[_tcslen(szDirectory) - 1] == TCHAR('\\'))			
			szDirectory[_tcslen(szDirectory) - 1] = TCHAR('\0');		
		LPSTR lpszLastDir = _tcsrchr(szDirectory, TCHAR('\\'));		
		if(lpszLastDir == NULL)			
			return true;		
		ZeroMemory(szShorterPath, MAX_PATH);		
		_tcsncpy(szShorterPath, szDirectory, _tcslen(szDirectory) - _tcslen(lpszLastDir));	
	}else		
		return true;		
	
	if(CreateDirectorySeq(szShorterPath))		
		return (CreateDirectory(szDirectory, NULL) == TRUE);		

	return false;
}
