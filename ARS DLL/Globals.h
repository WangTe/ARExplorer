// Global functions for logging

#ifndef _____1923089_GLOBALS_H____
#define _____1923089_GLOBALS_H____

#ifndef CTIME_TEXT
// Defines what format to use when displaying CTime values as text
#define CTIME_TEXT "%a %b %d, %Y %I:%M:%S %p"
#endif

class AFX_EXT_CLASS CFileLogging {
	// Global variables declared in Globals.cpp
	CStdioFile *pLogFile;
	int iLogLevel;

public:
	enum log_levels{level_dll = 1, level_lower = 2};
	CFileLogging();
	void Enable(CStdioFile *pNewLogFile, int iNewLogLevel = 1);
	void Disable();
	void ChangeLevel(int iNewLogLevel);
	void Log(int iLevelCheck, const char *pText, ...);

};

extern CFileLogging log;

#endif