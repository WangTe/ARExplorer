#ifndef __DAYS_OF_MONTH_H
#define __DAYS_OF_MONTH_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <afxtempl.h>

/////////////////////////////////////////////////////////////////
// structures and typdefs for Schedule contained in a backup job
class AFX_EXT_CLASS CDaysOfWeek : public CArray<UINT, UINT*> {
DECLARE_SERIAL(CDaysOfWeek);
public:
	enum days {
		sunday = 1,
		monday = 2,
		tuesday = 3,
		wednesday = 4,
		thursday = 5,
		friday = 6,
		saturday = 7
	};
};

#endif