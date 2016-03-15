
#include "StdAfx.h"

#include "Daily.h"

#include <math.h>

IMPLEMENT_SERIAL(CDaily, CObject, 1);

void CDaily::AdvanceNextRun(CTime &t, CTime &StartTime)
{
	// if the jobs next run time, is before t, advance it to next run interval
	//while( NextRun.GetTime() < t.GetTime() )
	//	NextRun = time_t (NextRun.GetTime() * CDaily::day_seconds);

//	double d = (t.GetTime() - NextRun.GetTime()) / (uiInterval * CDaily::day_seconds);
//	NextRun = NextRun + (time_t)ceil(d);

	// Check current date with start time
	NextRun = CTime(t.GetYear(), 
					  t.GetMonth(), 
					  t.GetDay(), 
					  StartTime.GetHour(), 
					  StartTime.GetMinute(), 
					  StartTime.GetSecond());
	if(t <= NextRun) {
		// The next run time hasn't elapsed today, so just return
		return;
	}else {
		// The next run time has elapsed for today, add 24 hours and return
		NextRun += CDaily::day_seconds;
		return;
	}
}

CDaily::CDaily()
{
	NextRun = 0;
	uiInterval = 0;
}


CDaily & CDaily::operator =(CDaily &d)
{
	uiInterval = d.uiInterval;
	NextRun = d.NextRun;
	return *this;
}
void CDaily::Serialize(CArchive &archive)
{
	CObject::Serialize( archive );
	if(archive.IsStoring()) { // Serialize
//		archive << StartTime;
		archive << uiInterval;
		archive << NextRun;
	}else { // De-serialize
//		archive >> StartTime;
		archive >> uiInterval;
		archive >> NextRun;
	}
}
