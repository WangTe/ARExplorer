#ifndef __WEEKLY_H
#define __WEEKLY_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


/////////////////////////////////////////////////////////////////
// CWeekly class
class AFX_EXT_CLASS CWeekly : public CObject {
DECLARE_SERIAL(CWeekly);
public:
	BOOL GetDay(int indexDay);
	BOOL GetHour(int indexHour);
	void SetHour(int indexHour, BOOL bValue);
	void SetDays(int iDay, BOOL bValue);
	void SetNextRunTime(CTime &newTime);
	CTime GetCurrentRunTime();
	CTime GetNextRunTime(CTime &StartTime);
	enum class_globals {week_seconds = 604800, /*seconds in a week*/
						hour_seconds = 3600, /*seconds in an hour*/
						num_days = 8, /*number of days in array*/
						min_days_index = 1,
						num_hours = 24, /*num hours in a day*/ 
						max_hours_index = 23 };
	CWeekly();
	CTime AdvanceNextRun(CTime &StartTime, BOOL bAdvance = TRUE);
	void Serialize( CArchive &archive );
	CWeekly & operator =(CWeekly &w);
//	UINT uiInterval; // number of days
//	CTime StartTime; // time of day
//	int	iMinute; // minute of the hour to execute on ex: 30 for 12:30, 1:30, 2:30, etc
protected:
	CTime NextRun;
	BOOL Days[num_days]; // 0=unused, 1=sunday, 2=monday, 1-7 Sun - Sat
	BOOL bHoursInDay[num_hours]; // 0 is 12 am, 23 is 11 pm
};

#endif