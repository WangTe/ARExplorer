#include "StdAfx.h"

// Include local file here (after StdAfx.h and before Schedule.h)
//#include "main.h"
#include "Daily.h"
#include "Weekly.h"
#include "Monthly.h"
#include "Schedule.h"
#include "Globals.h"

IMPLEMENT_SERIAL(CSchedule, CObject, 1);

CSchedule::CSchedule(CSchedule &s)
{
	operator=(s);
}

CSchedule::CSchedule()
{
	Once = 0;
	uiType = CSchedule::none;
	m_tStartTime = 0;

}

//////////////////////////////////////////////////////////////////////////////////
// Description: Gets next run time based on StartTime and current time
//				Doesn't save next run time in Weekly object
// Returns: CTime - the next run time
CTime CSchedule::GetNextRunTime()
{
	switch(uiType) {
	// Once Schedule
	case CSchedule::once:
	case CSchedule::now:		// Added this line: 8/18/08
		return Once.GetTime();
		break;
//	case CSchedule::daily:
//		Daily.AdvanceNextRun(t, m_tStartTime);		// advances to next valid run time
//		return Daily.NextRun.GetTime(); // return the next valid run time
//		break;
	// Weekly schedule
	case CSchedule::weekly:
		return Weekly.GetNextRunTime(m_tStartTime).GetTime(); // return the next valid run time
		break;
//	case CSchedule::monthly:
//		Monthly.AdvanceNextRun(t, m_tStartTime);
//		return Monthly.NextRun.GetTime();
//		break;
	}
	return 0;
}

CSchedule & CSchedule::operator =(CSchedule &s)
{
	uiType = s.uiType;
	Once = s.Once;
	Daily = s.Daily;
	Weekly = s.Weekly;
	Monthly = s.Monthly;
	m_tStartTime = s.m_tStartTime;
	return *this;
}

void CSchedule::Serialize(CArchive &ar)
{
	CObject::Serialize( ar );

	if(ar.IsStoring()) { // Serialize - save to archive
		ar << uiType;
		ar << Once;
		ar << m_tStartTime;
	}else { // Deserialize - Restore from archive
		ar >> uiType;
		ar >> Once;
		ar >> m_tStartTime;
	}
	Daily.Serialize( ar );
	Weekly.Serialize( ar );
	Monthly.Serialize( ar );

	log.Log(CFileLogging::level_lower, "CSchedule:Serialize(): Done serializing schedule.", NULL);
}

//////////////////////////////////////////////////////////////////////
// Description: returns the type of the schedule.  Possible types are:
//		once,
//		daily,
//		weekly,
//		monthly,
//		now
UINT CSchedule::GetType()
{
	return uiType;
}

CTime CSchedule::GetStartTime()
{
	return m_tStartTime;
}


//DEL CTime CSchedule::GetLastRunTime()
//DEL {
//DEL 	return m_tLastRunTime;
//DEL }

CTime & CSchedule::SetStartTime(CTime &m_Time)
{
	// store the start time
	m_tStartTime = m_Time; 

	return m_tStartTime;
} // end of SetStartTime()

////////////////////////////////////////////////////////////////////////
// Description: If Once job, returns Once time
//				If Weekly job, return next run time.  Saves next run time
//				if bAdvance is TRUE
// Returns: On success, returns either Once time or NextRun time.
//			On error, returns CTime set to Jan 1, 1973 12:00:00 AM
CTime CSchedule::AdvanceNextRun(BOOL bAdvance)
{
	switch(uiType) {
	// Once can't advance, return Once time
	case CSchedule::once:
		return Once;
		break;
	// Advance/Get weekly next run time
	case CSchedule::weekly:
		return Weekly.AdvanceNextRun(m_tStartTime, bAdvance);
		break;
	}

	// shouldn't get here, return 0 indicating error
	return CTime(0);
}

void CSchedule::SetType(unsigned int uiNewType)
{
	uiType = uiNewType;
}

CTime CSchedule::GetCurrentRunTime()
{
	switch(uiType) {
	// Once can't advance, return Once time
	case CSchedule::once:
		return Once;
		break;
	// Advance/Get weekly next run time
	case CSchedule::weekly:
		return Weekly.GetCurrentRunTime();
		break;
	}

	// shouldn't get here, return 0 indicating error
	return CTime(0);
}

//DEL void CSchedule::SetLastRunTime(CTime &LastRun)
//DEL {
//DEL 	m_tLastRunTime = LastRun;
//DEL }

/////////////////////////////////////////////////////////////////////////////////
// Description: USE WITH EXTREME CAUTION! if used wrong, could mess up scheudling
//				Sets the NextRun time. Should only be set by a CTime gotten 
//				from an immediately previous call to GetNextRunTime()
void CSchedule::SetNextRunTime(CTime &newTime)
{
	if(uiType == CSchedule::weekly)
		Weekly.SetNextRunTime(newTime);
}

void CSchedule::SetWeeklyDays(int indexDay, BOOL bValue)
{
	Weekly.SetDays(indexDay, bValue);
}

void CSchedule::SetWeeklyHour(int indexHour, BOOL bValue)
{
	Weekly.SetHour(indexHour, bValue);
}

void CSchedule::SetOnceTime(CTime newTime)
{
	Once = newTime;
}

CTime CSchedule::GetOnceTime()
{
	return Once;
}

BOOL CSchedule::GetWeeklyHour(int indexHour)
{
	return Weekly.GetHour(indexHour);
}

BOOL CSchedule::GetWeeklyDay(int indexDay)
{
	return Weekly.GetDay(indexDay);
}
