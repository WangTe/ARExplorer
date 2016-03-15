#ifndef __MONTHLY_H
#define __MONTHLY_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "DaysOfMonth.h"

/////////////////////////////////////////////////////////////////
// CMonthly class
class AFX_EXT_CLASS CMonthly : public CObject {
DECLARE_SERIAL(CMonthly);
public:
	int iDay;
	void AdvanceNextRun(CTime &t, CTime &StartTime);
	CTime NextRun;
	CMonthly();
	void Serialize( CArchive &ar );
	CMonthly & operator =(CMonthly &m);
//	CTime StartTime; // time of day
//	CDaysOfWeek Days; // day of month to run job
};


#endif