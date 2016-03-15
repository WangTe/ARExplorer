#ifndef __SCHEDULE_H
#define __SCHEDULE_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "Daily.h"
#include "Weekly.h"
#include "Monthly.h"

/////////////////////////////////////////////////////////////////
// CSchedule class
class AFX_EXT_CLASS CSchedule : public CObject {
DECLARE_SERIAL(CSchedule);
public:
	BOOL GetWeeklyDay(int indexDay);
	BOOL GetWeeklyHour(int indexHour);
	CTime GetOnceTime();
	void SetOnceTime(CTime newTime);
	void SetWeeklyHour(int indexHour, BOOL bValue);
	void SetWeeklyDays(int indexDay, BOOL bValue);
	void SetNextRunTime(CTime &newTime);
	CTime GetCurrentRunTime();
	void SetType(unsigned int uiNewType);
	CTime AdvanceNextRun(BOOL bAdvance = TRUE);
	CTime & SetStartTime(CTime &m_Time);
	CTime GetStartTime();
	UINT GetType();
	enum Types {
		none = 0,
		once,
		daily,
		weekly,
		monthly,
		now
	};
	CSchedule();
	CSchedule(CSchedule &s);
	void Serialize( CArchive &ar );
	CSchedule & operator =(CSchedule &s);
	CTime GetNextRunTime();
protected:
	CTime Once;
	CWeekly Weekly;
	CDaily Daily;
	CMonthly Monthly;
	unsigned int uiType;
	CTime m_tStartTime;
};


#endif