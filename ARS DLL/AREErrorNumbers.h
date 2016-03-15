#ifndef	ARE_ERROR_NUMBERS_H
#define ARE_ERROR_NUMBERS_H


//////////////////////////////////////////////////////////////////////////////////////////////////
// NOTE:	For all custom error messages, the class which does the error checkig must			//
//			do two things:																		//
//			1) Populate strErrorText with the error text.										//
//			2) Store the error number in iLastErrorNum.											//
//			3) return iLastErrorNum.															//
//////////////////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////////
// All possible error message for all classes													//
#define MSG_NO_USER				"ERROR(10500): There is no user specified."
#define MSG_NO_SERVER			"ERROR(10501): There is no server specified."
#define MSG_TOO_MANY_FORMS		"ERROR(10502): Retrieved too many forms."	// CFormList couldn't hold all of the forms returned by the server.  You need to make CFormList handle the list of forms dynamically.
#define MSG_INVALID_REL_OP		"ERROR(10503): Encountered an unhandled type of \"ARFieldValueOrArithStruct\" structure."  // Used by QREQueryStoreRelValue() function
#define MSG_UNKNOWN				"ERROR(10504): An unknown error has occurred.  Please contact AR Accelerators support."
#define MSG_UNKNOWN_ARITH_TYPE	"ERROR(10505): Decoding an ARQualifierStruct failed when an unkown arithmetic type was encountered."	// Used by AREQueryStoreRelValue() function.
#define MSG_NO_FORMS			"ERROR(10506): There were no forms specified."	// Used in CSearchAL::PopulateObjectList()
#define MSG_END_NAME_LIST		"ERROR(10507): Ran out of storage space for CMyNameList."
#define MSG_UNKNOWN_OPERATOR	"ERROR(10508): Unknown operator encountered in Run If line."		// Used by QueryStoreRelOp()
#define MSG_STAT_HIST_VALUE		"ERROR(10509): Status history was found in relational value.  These are not supported yet."		// Used by QueryStoreRelValue()
#define MSG_QUERY_VALUE			"ERROR(10510): Query value found in relational value.  These aren't supported yet."		// Used by QueryStoreRelValue()
#define MSG_UNKNOWN_VAL_STRUCT	"ERROR(10511): Unknown value structure found in Run If line."
#define MSG_WRONG_OPERATOR		"ERROR(10512): The operator you have specified isn't supported."
#define MSG_OUT_OF_RANGE		"ERROR(10513): The number specified was out of the allowed range."
#define MSG_ENABLED_DISABLED	"ERROR(10514): Only Enabled or Disabled is allowed as the parameter."
#define MSG_OLD_INTERFACE		"ERROR(10515): Internal Error.  This interface is no longer supported.  Please contact AR Accelerators support."
#define MSG_FIELD_TOO_LONG		"ERROR(10516): A field name couldn't be stored because it was too long."
#define MSG_NOT_ENOUGH_MEMORY	"ERROR(10517): Memory allocation failed.  There was not enough virtual memory available.  Close some applicaitons and try again."
#define MSG_NO_FIELD_ID			"ERROR(10518): The field id wasn't specified."
#define MSG_NO_BITMASK			"ERROR(10519): No bitmask was supplied.  Cannot search the Execute On parameter."
#define MSG_BITMASK_TOO_BIG		"ERROR(10520): The bitmask which have been specified is too big.  Please contact AR Accelerators support."
#define MSG_NAME_TOO_LONG		"ERROR(10521): Internal Error.  The name specified in CLinkedList is too long.  Please contact AR Accelerators support."
#define MSG_NULL_POINTER_PARAM	"ERROR(10522): Internal Error.  A NULL pointer was passed as a parameter.  Please contact AR Accelerators support."
#define MSG_TOO_MANY_FILTERS	"ERROR(10523): More than one Filter cannot be specified when changing the Filter Name."
#define MSG_NO_FILTERS			"ERROR(10524): No Filters were selected."
#define MSG_UNKNOWN_MOD_FILTER	"ERROR(10525): Internal Error: Please contact AR Accelerators support with error code 10525."
#define MSG_NO_TEXT				"ERROR(10526): Internal Error: No text was specified.  Please contact AR Accelerators support."
#define MSG_NO_AL_NAME			"ERROR(10527): Internal Error: No Active Link name was specified for retrieval.  Please contact AR Accelerators support."
#define MSG_NO_GROUP_NAMES		"ERROR(10528): Internal Error: No groups were given to update an object.  Please contact AR Accelerators support."
#define MSG_TOO_MANY_AL			"ERROR(10529): More than one Active Link cannot be specified when changing the Active Link name."
#define MSG_NOT_APPEND_PERM_ACTION "ERROR(10530): Internal Error: Only the append permissions action is allowed.  Please contact AR Accelerators support."
#define MSG_WRONG_PERM_ACTION	"ERROR(10531): Internal Error: The permission modification action specified was incorrect.  Please contact AR Accelerators support."
#define MSG_NO_FL_NAME			"ERROR(10532): Internal Error: No Filter Name was specfified for retrieval."
#define MSG_NO_FINISHED_ID		"ERROR(10533): Internal Error: No Finished Id list was available.  Please contact AR Accelerators support."
#define MSG_NO_ACTION_LIST		"ERROR(10534): Internal Error: The action list wasn't populated."
#define MSG_NOT_ACTION_MESSAGE	"ERROR(10535): Internal Error: The action type isn't Message."

#define MSG_NOT_REGISTERED		"ERROR(10999): You are attempting to access the COM object directly.  This isn't allowed, you are violating AR Accelerators copyright on this product."


#define MSG_NULL_POINT_FIELDS	"ERROR(11001): A null pointer assignment was found in a CFieldList object."



//////////////////////////////////////////////////////////////////////////////////////////////////
// All possible error numbers for all class														//
#define ERR_NO_USER				10500		// Used by CLoginInfo
#define ERR_NO_SERVER			10501		// Used by CLoginInfo
#define ERR_TOO_MANY_FORMS		10502		// Used by CFormList
#define ERR_INVALID_REL_OP		10503		// Used by QREQueryStoreRelValue() function
#define ERR_UNKNOWN				10504		// Used everywhere
#define ERR_UNKNOWN_ARITH_TYPE	10505		// Used by AREQueryStoreRelValue() function.
#define ERR_NO_FORMS			10506		// Used in CSearchAL::PopulateObjectList()
#define ERR_END_NAME_LIST		10507		// Used by CMyNameList::AddNameToList()
#define ERR_UNKNOWN_OPERATOR	10508		// Used by QueryStoreRelOp()
#define ERR_STAT_HIST_VALUE		10509		// Used by QueryStoreRelValue()
#define ERR_QUERY_VALUE			10510		// Used by QueryStoreRelValue()
#define ERR_UNKNOWN_VAL_STRUCT	10511		// Used by QueryStoreValueStruct()
#define ERR_WRONG_OPERATOR		10512		// Used by CSearchAL::SetModifiedTime(), indicates a greater or less than operator wasn't passed as parameter 2.
#define ERR_OUT_OF_RANGE		10513		// Used by CSearchAL::ParamExecutionOrder(), indicates the specified execution order was out of the range 0 to 1000.
#define ERR_ENABLED_DISABLED	10514		// Used by CSearchAL::ParamEnabled()
#define ERR_OLD_INTERFACE		10515		// Used by the interface, means function not supported anymore.
#define ERR_FIELD_TOO_LONG		10516		// Used by the CFieldList class to indicate a field name returned by ARS was over 30 chars.
#define ERR_NOT_ENOUGH_MEMORY	10517		// Can be used by any class which dynamically allocates memory
#define ERR_NO_FIELD_ID			10518		// Used in CUpdateFieldProps 
#define ERR_NO_BITMASK			10519		// Used in CSearchFL::ParamExecuteOn and CSearchAL::ParamExecuteOn
#define ERR_BITMASK_TOO_BIG		10520		// Used in CExecuteOnAL and CExecuteOnFL
#define ERR_NAME_TOO_LONG		10521		// Used in CLinkedList::AddNameToList()
#define ERR_NULL_POINTER_PARAM	10522		// Indicates a NULL pointer used as a paramter
#define ERR_TOO_MANY_FILTERS	10523		// Only 1 filter may be specified when modifying the filter name property.
#define ERR_NO_FILTERS			10524		// No filters were specified to be modified.
#define ERR_UNKNOWN_MOD_FILTER	10525		// An unknown error occurred in CModifyFilter class
#define ERR_NO_TEXT				10526		// Indicates no text was passed as a needed parameter.
#define ERR_NO_AL_NAME			10527		// No active link name was specified before calling ARGetActiveLink() or ARSetActiveLink()
#define ERR_NO_GROUP_NAMES		10528		// No groups were listed when CUpdateActiveLink::CopyPermGroupToStruct() was called
#define ERR_TOO_MANY_AL			10529		// Only 1 active link may be specified when modifying the active link name property.
#define ERR_NOT_APPEND_PERM_ACTION 10530	// The only action allowed is an APPEND PERMISSIONS action.
#define ERR_WRONG_PERM_ACTION	10531		// The Permission action specified was incorrect.
#define ERR_NO_FL_NAME			10532		// No filter name was given before calling CGetFLProperties::StoreHelpText()
#define ERR_NO_FINISHED_ID		10533		// Used in CUpdateActiveLink::CopyPermGroupToStructAppend(), means the Finished Group Id list was blank for some reason
#define ERR_NO_ACTION_LIST		10534		// The ActionList wasn't populated in a Active Link, Filter or Escalation object.
#define ERR_NOT_ACTION_MESSAGE	10535		// The action type wasn't a message type.

#define ERR_NOT_REGISTERED		10999		// The DLL hasn't been registered.  Invalid access attempt to COM object.

#define ERR_NULL_POINT_FIELDS	11001		// Used in CFieldList only to indicate a NULL pointer assignment

//////////////////////////////////////////////////////////////////////////////////////////////////
// All possible warnings numbers...																		//
#define WARN_END_FIELDS_LIST		20001	// Used by CFieldList class to indicate at end of MAX_FIELDS.
#define WARN_UNKNOWN_ACTION_TYPE	20002

//////////////////////////////////////////////////////////////////////////////////////////////////
// All possible warning messages...
#define WMSG_END_FIELDS_LIST		"WARNING(20001): You have exceded the maximum number of fields for a CFieldList object."
#define WMSG_UNKNOWN_ACTION_TYPE	"WARNING(20002): An unknown action type has been encountered." // An unknown action type usually means the user is running a higher version of ARS.


//////////////////////////////////////////////////////////////////////////////////////////////////
// All possible information numbers...
#define INUM_WRONG_ENTRY_MODE	30001

//////////////////////////////////////////////////////////////////////////////////////////////////
// All possible information messages...
#define INFO_WRONG_ENTRY_MODE	"Note 30001: The field's Entry Mode couldn't be updated because the field is a Core field or Display Only field."





#endif	// End of file.