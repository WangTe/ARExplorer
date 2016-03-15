#ifndef __DAILY_H
#define __DAILY_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


/////////////////////////////////////////////////////////////////
// CDaily class
class AFX_EXT_CLASS CDaily : public CObject {
DECLARE_SERIAL(CDaily);
public:
	CDaily();
	enum seconds_in { day_seconds = 86400, /*seconds in a day*/ 
					  secs_to_nanosecs = 10000000 };
	void AdvanceNextRun(CTime &t, CTime &StartTime);
	CTime NextRun;
	void Serialize( CArchive &archive );
	CDaily & operator =(CDaily &d);
	UINT uiInterval; // number of days
//	CTime StartTime; // time of day
};

#endif
