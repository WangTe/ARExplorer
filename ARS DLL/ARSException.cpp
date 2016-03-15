// ARSException.cpp: implementation of the CARSException class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ARSException.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif


IMPLEMENT_DYNAMIC(CARSException,CException);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CARSException::CARSException(ARStatusList &pStatusList)
{
	int             i;         /* working index */
	ARStatusStruct *pStatusStruct;   /* working pointer */
	char           tempString[31];  // Char array

	strErrorText.Empty();
	try{
		// If there are items in the status list then store them
		if (pStatusList.numItems != 0)
		{
			// Store the address of the status list.
			pStatusStruct = pStatusList.statusList;

			// Store the 1st error number in iLastErrorNum
			iErrorNum = pStatusStruct->messageNum;

			// Process each item in the status list
			for (i = 0; i < (int) pStatusList.numItems; i++)
			{
				if(i > 0)
				{
					// There is more than one message, insert carriage returns.
					strErrorText += "\n\n";
				}	// End if

				switch (pStatusStruct->messageType)
				{
				case AR_RETURN_OK:
					strErrorText += "NOTE(";
					uiType = AR_RETURN_OK;
					break;
				case AR_RETURN_WARNING:
					strErrorText += "WARNING(";
					uiType = AR_RETURN_WARNING;
					break;
				case AR_RETURN_ERROR:
					strErrorText += "ERROR(";
					uiType = AR_RETURN_ERROR;
					break;
				default:
					strErrorText += "UKNOWN(";
					uiType = 10;
					break;
				}	// End of switch ARStatusList

				// Convert and store the message number in the error message text.
				_itoa( pStatusStruct->messageNum, tempString, 10 );
				strErrorText += tempString;
				strErrorText += "): ";
         
				// Store the long text of the message.
				if (pStatusStruct->messageText == NULL) 
				{
					// There was no error text, terminate the error message.
					strErrorText += "): No error text available";
				}
				else 
				{
					// There was some error text, store it in strErrorText
					strErrorText += pStatusStruct->messageText;
				}
				
				pStatusStruct++;
			}		// End of for loop
			strErrorText += "\0";

		}	// End of if
	}
	catch(CARSException *)
	{
		FreeARStatusList(&pStatusList, FALSE);
		throw;
	}
	// Free the status list
	FreeARStatusList(&pStatusList, FALSE);
}

CARSException::~CARSException()
{

}

CARSException::CARSException()
{
	uiType = 0;
	iErrorNum = 0;
}

CARSException::CARSException(const char *pText)
{
	iErrorNum = 10000; // Generic number for custom errors.
	strErrorText = pText;
}

///////////////////////////////////////////////////////////////////////////
// Throws an ARS Exception and free the status list memory
void ThrowARSException(ARStatusList &StatusList, const char *pFunctionName)
{
	CARSException *pError = new CARSException(StatusList);
	pError->strErrorText += "\n";
	pError->strErrorText += pFunctionName;
	FreeARStatusList(&StatusList, FALSE);
	throw pError;
}

void ThrowARSException(const char *cpText, const char *pFunctionName)
{
	CARSException *pError = new CARSException(cpText);
	pError->strErrorText += "\n";
	pError->strErrorText += pFunctionName;
	throw pError;
}

//void Log(CString strOutputText)
//{
	// insert code to output strOutputText to error logging file.
	// you probably want to use a callback function so the gui
	// can register the function to call for outputing to error log
//} 

void Log(const char *pOutputText)
{
	CString strOutputText(pOutputText);
	Log(strOutputText);
}
