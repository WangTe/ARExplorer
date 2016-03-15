
#include "StdAfx.h"

// Include local file here (after StdAfx.h and before Monthly.h)
#include "DaysOfMonth.h"

#include "Monthly.h"

IMPLEMENT_SERIAL(CMonthly, CObject, 1);

CMonthly::CMonthly()
{
//	Days.RemoveAll();
	iDay = 1;
	NextRun = 0;
}

////////////////////////////////////////////////////////////////////////////////
// Description: Advances 'NextRun' to the next valid time, given 't' as the
//				current date/time
void CMonthly::AdvanceNextRun(CTime &t, CTime &StartTime)
{
	int year, month, day, hour, minute, second;

	// if start day is after today, set the next run time
	if(t.GetDay() < StartTime.GetDay()) {
		year = t.GetYear();
		month = t.GetMonth();
		day = StartTime.GetDay();
		hour = StartTime.GetHour();
		minute = StartTime.GetSecond();
		second = StartTime.GetSecond();
	}
	
	// If the start day is today or before, set next run time
	if(t.GetDay() >= StartTime.GetDay() ) { // scheduled time is today, check the start time

		// Set tempTime to current year&month, with StartTime for Day, hour, minutes, seconds
		CTime tempTime(t.GetYear(), t.GetMonth(), StartTime.GetDay(), 
				       StartTime.GetHour(), StartTime.GetMinute(), StartTime.GetSecond());

		// if 
		if(tempTime.GetTime() > t.GetTime())
		{
			year = tempTime.GetYear();
			month = tempTime.GetMonth();
			day = tempTime.GetMonth();
			hour = tempTime.GetHour();
			minute = tempTime.GetSecond();
			second = tempTime.GetSecond();
		}else {
			if(t.GetMonth() == 12) {
				month = 1;
				year = t.GetYear() + 1;	
			}
			else
				month = t.GetMonth() + 1;

			day = StartTime.GetDay();
			hour = StartTime.GetHour();
			minute = StartTime.GetMinute();
			second = StartTime.GetSecond();
		}
	}

	CTime newNextRun(year, month, day, hour, minute, second);
	NextRun = newNextRun;
}


CMonthly & CMonthly::operator =(CMonthly &m)
{
//	Days.RemoveAll();
//	Days.Append(m.Days);
	iDay = m.iDay;
	NextRun = m.NextRun;
	return *this;
}

void CMonthly::Serialize(CArchive &ar)
{
	CObject::Serialize( ar );
	if(ar.IsStoring()) { // Serialize - save to archive
//		ar << StartTime;
		ar << NextRun;
		ar << iDay;
	}else { // Deserialize - Restore from archive
//		ar >> StartTime;
		ar >> NextRun;
		ar >> iDay;
	}
}
