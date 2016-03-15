// Global functions for logging

#include "StdAfx.h"
#include "Globals.h"

/////////////////////////////////////////////////////////////////////////
// Description: Enables logging in ARS DLL module.  Must have a valid
//				pointer to a CStdioFile and valid level 1 or 2
void CFileLogging::Enable(CStdioFile *pNewLogFile, int iNewLogLevel) {
	CFileStatus m_FileStatus;

	// If the file is valid, store the file & log level information
	if(pNewLogFile && pNewLogFile->GetStatus(m_FileStatus)) {
		pLogFile = pNewLogFile;
		if(iNewLogLevel > 0 && iNewLogLevel <= 3)
			iLogLevel = iNewLogLevel;
	}
}

////////////////////////////////////////////////////////////////////////
// Description: Disables logging in ARS DLL module
void CFileLogging::Disable() {
	pLogFile = NULL;
}

////////////////////////////////////////////////////////////////////////
// Description: Sets the level of logging
//				1 = ARS DLL level logging
//				2 = lower level logging like CBlockingSocket, etc.
//				Note: only 1-2 are valid
void CFileLogging::ChangeLevel(int iNewLogLevel) {
	if(iNewLogLevel > 0 && iNewLogLevel <= 2)
		iLogLevel = iNewLogLevel;
}

////////////////////////////////////////////////////////////////////////
// Description: OUtputs to the log file (if enabled) and flushes stream
void CFileLogging::Log(int iLevelCheck, const char *pText, ...) {
	// Check for logging enabled, exit if it isn't enabled
	if(!pLogFile)
		return;
	if(iLogLevel < iLevelCheck)
		return;

	char *pNextText = NULL;

	// Init the list of arguments
	va_list list;
	va_start(list,pText);

	for(int i=0;;i++) {
		pNextText=va_arg(list, char *);
		if(pNextText == NULL)
			break;

		// Output formated date time, at beginning of line
		if(i==0) {
			// Return the current date using standard format
			CTime time = CTime::GetCurrentTime();
			CString strTime;
			strTime = time.Format(CTIME_TEXT); // create time/date using proper format
			strTime += ":\t"; // add formatting
			pLogFile->WriteString(LPCTSTR(strTime));
		}

		// Output the text to the log file
		pLogFile->WriteString(pNextText);
	}

	// cleanup the list
	va_end(list);

	// Flush the file output
	pLogFile->Flush();
}

CFileLogging::CFileLogging()
{
	pLogFile = NULL;
	iLogLevel = 0;
}
