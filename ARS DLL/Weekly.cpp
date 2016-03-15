#include "StdAfx.h"

#include "DaysOfMonth.h"

#include "Weekly.h"

#include <math.h>

IMPLEMENT_SERIAL(CWeekly, CObject, 1);
IMPLEMENT_SERIAL(CDaysOfWeek, CObject, 1);

CWeekly::CWeekly()
{
//	Days.RemoveAll();
	for(int i=0; i<num_days; i++)
		Days[i] = 0;
	for(i=0; i<num_hours; i++)
		bHoursInDay[i] = 0;
	NextRun = 0;
//	uiInterval = 0;
}

CWeekly & CWeekly::operator =(CWeekly &w)
{
	// save Days
	for(int i=0; i<num_days; i++) 
		Days[i] = w.Days[i];

	// Save bHoursInDay
	for(i=0; i<num_hours; i++) 
		bHoursInDay[i] = w.bHoursInDay[i];

	// Save NextRun
	NextRun = w.NextRun;

	return *this;
}

////////////////////////////////////////////////////////////////////////////////
// Description: Advances 'NextRun' to the next valid time, given 't' as the
//				current date/time.  Only advances saves next run date in 
//				NextRun if bAdvance is TRUE
// Returns:	CTime - the next run time
CTime CWeekly::AdvanceNextRun(CTime &StartTime, BOOL bAdvance)
{
	CTime t = CTime::GetCurrentTime();

	// Only advance next run if bAdvance is set to TRUE
	CTime tempNextRun;
	CTime *pNextRun;

	if(bAdvance)
		pNextRun = &NextRun;
	else
		pNextRun = &tempNextRun;

	// indexes for checking bHoursInDay and Days arrays
	int iCurrWeekday; // current weekday (1-8) Sun-Sat
	int iCurrHour; // current hour (0-23) 12am - 11pm
	
//	int iCurrDay; // current day in month (1-31)
//	int iCurrMonth; // current month (1-12) Jan-Dec
//	int iCurrYear; // current year (2004-->)

	int iNumLoops = 7 * 24; // num days in week, times num hours in day == total number times to loop


	CTimeSpan tsOneHour(hour_seconds); // timespan for one hour


	// Initialize time to start checking from (tStartCheckTime)
	if(t.GetTime() > StartTime.GetTime()) {
		// now is after StartTime, use now
		*pNextRun = CTime(t.GetYear(), 
						  t.GetMonth(),
						  t.GetDay(),
						  t.GetHour(),
						  NextRun.GetMinute(),
						  NextRun.GetSecond() ); // bug fixed, should use hour & second from NextRun, but hour & m/d/y from t
		iCurrWeekday = t.GetDayOfWeek();
		iCurrHour = t.GetHour();

		// ensure pNextRun isn't set to date/time in the past, this will happen if
		// Start Time is set to current hour, but minutes have already past
		if(t.GetTime() > pNextRun->GetTime() ) {
			*pNextRun = *pNextRun + tsOneHour;
			iCurrWeekday = pNextRun->GetDayOfWeek();
			iCurrHour = pNextRun->GetHour();
		}
	}else {
		// StartTime is after now, use StartTime
		*pNextRun = StartTime;
		iCurrWeekday = StartTime.GetDay();
		iCurrHour = StartTime.GetHour();
	}

	// will loop once for every hour in every day
	for(int i=0; i<iNumLoops; i++) {

		if(Days[iCurrWeekday] == TRUE) {
			// the current day was selected, check the current hour
			if(bHoursInDay[iCurrHour] == TRUE) {
			//	*pNextRun = *pNextRun + tsOneHour;
				return *pNextRun;
			}
		}

		// advance to next hours and days index
		if(iCurrHour < max_hours_index)
			iCurrHour++;
		else {
			iCurrHour = 0;
			// advamce to next day index
			if(iCurrWeekday < num_days)
				iCurrWeekday++;
			else
				iCurrWeekday = min_days_index; // starts back at Sundays index
		}

		// add one hour of time
		*pNextRun = *pNextRun + tsOneHour;
	}

	ASSERT(TRUE);
	return CTime(0);

//	// init NextRun if needed
//	if(NextRun.GetTime() == 0)
//		NextRun = CTime::GetCurrentTime();
//
//	ASSERT(NextRun.GetTime() < t.GetTime());
//
//	// Find next valid run time
//	double d_SkipAmount(uiInterval * CWeekly::week_seconds);
//	double d_SkipTimes(    ( t.GetTime() - NextRun.GetTime() ) / d_SkipAmount    );
//	CTimeSpan tsDay(1, 0, 0, 0); // timespan of 1 day
//
//	d_SkipTimes = ceil(d_SkipTimes); 
//	NextRun = NextRun.GetTime() + (time_t)(d_SkipTimes * d_SkipAmount); // puts NextRun in middle of next run week
//
//	// Backup until we're at the start of the week to run job on.
//	while(NextRun.GetDayOfWeek() != CDaysOfWeek::sunday)
//		NextRun-= tsDay;
//
//	CTime newRunTime(NextRun.GetYear(),		// year
//					 NextRun.GetMonth(),	// month
//					 NextRun.GetDay(),		// day
//					 0,						// hour
//					 0,						// minute
//					 0);					// second
//
//	for(int i=0; i < 8; i++) {
//		// if current day is selected is Days, then add in that start time and save NextRun time.
//		if(Days[i] == TRUE) {
///			CTimeSpan addTime(0, StartTime.GetHour(), StartTime.GetMinute(), StartTime.GetSecond() );
//			newRunTime += addTime;
//			NextRun = newRunTime;
//			break;
//		}
//		// else, advance newRunTime to next day at midnight
//		else {
//			newRunTime += tsDay;
//		}
//	} // end for()
	
} // end AdvanceNextRun()

void CWeekly::Serialize(CArchive &archive)
{
	CObject::Serialize( archive );
	if(archive.IsStoring()) { // Serialize - save to archive
//		archive << StartTime;
//		archive << uiInterval;
		archive << NextRun;
		for(int i=0; i<num_days; i++)
			archive << Days[i];
		for(int x=0; x<num_hours; x++)
			archive << bHoursInDay[x];
	}else { // Deserialize - Restore from archive
//		archive >> StartTime;
//		archive >> uiInterval;
		archive >> NextRun;
		for(int i=0; i<num_days; i++)
			archive >> Days[i];
		for(int x=0; x<num_hours; x++)
			archive >> bHoursInDay[x];
	}
}

//////////////////////////////////////////////////////////////////////////////////
// Description: Gets next run time based on StartTime and current time
//				Doesn't save next run time in Weekly object
// Returns: CTime - the next run time
CTime CWeekly::GetNextRunTime(CTime &StartTime)
{
	return AdvanceNextRun(StartTime, FALSE);
}

CTime CWeekly::GetCurrentRunTime()
{
	return NextRun;
}

void CWeekly::SetNextRunTime(CTime &newTime)
{
	NextRun = newTime;
}

//////////////////////////////////////////////////////////////////////////////////////////
// Description: Sets the day of week for schedule ONLY if iDay is between 1-7.
void CWeekly::SetDays(int iDay, BOOL bValue)
{
	if(iDay > 0 && iDay < num_days) {
		Days[iDay] = bValue;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////
// Description: Sets the hours of day for schedule ONLY if indexHour is between 0-23
void CWeekly::SetHour(int indexHour, BOOL bValue)
{
	if(indexHour >= 0 && indexHour < num_hours) {
		bHoursInDay[indexHour] = bValue;
	}
}

//////////////////////////////////////////////////////////////////////////////////////////
// Description: Returns the hour value indicated by indexHour.  Return num_hours
//				if indexHour is out of bounds
BOOL CWeekly::GetHour(int indexHour)
{
	if(indexHour >= 0 && indexHour < num_hours) {
		return bHoursInDay[indexHour];
	}

	return num_hours; // return error
}

//////////////////////////////////////////////////////////////////////////////////////////
// Description: Returns the day value indicated by indexDay.  If indexDay is out of 
//				bounds returns num_days as an error code.
BOOL CWeekly::GetDay(int indexDay)
{
	if(indexDay >= 0 && indexDay < num_days)
		return Days[indexDay];

	return num_days; // return error
}
