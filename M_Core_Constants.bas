Attribute VB_Name = "M_Core_Constants"
Option Explicit
'*******************************************************************************
' Module:      M_Core_Constants
'
' Purpose:
'   Centralized constants for workbook metadata, sheet names, table names,
'   key column names, audit column names, log levels, and shared messages.
'   This module eliminates magic strings in code and keeps all naming aligned
'   with Schema 3.4.3 (TBL_SCHEMA on SCHEMA tab).
'
' Inputs (Tabs/Tables/Headers):
'   - None (this module does not read from the workbook).
'
' Outputs / Side effects:
'   - Public Const values used by all other modules for:
'       * Sheet (tab) names
'       * Table (ListObject) names
'       * Common column names (keys, audit fields, log fields)
'       * Application and schema versions
'       * Log levels and generic messages
'
' Preconditions:
'   - Schema 3.4.3 is the active, enforced schema in SCHEMA!TBL_SCHEMA.
'
' Postconditions:
'   - None at runtime; this is a static configuration module.
'
' Errors & Guards:
'   - None; this module contains only constants.
'
' Version:     v0.1.0
' Author:      ChatGPT (assistant)
' Date:        2025-11-28
'*******************************************************************************

'===============================================================================
' Application / Schema Versioning
'===============================================================================

Public Const APP_VERSION      As String = "0.1.0"
Public Const SCHEMA_VERSION   As String = "3.4.3"

'===============================================================================
' Sheet (Tab) Names  - must match workbook tabs exactly
'===============================================================================

Public Const SH_LANDING       As String = "Landing"
Public Const SH_SCHEMA        As String = "SCHEMA"
Public Const SH_AUTO          As String = "Auto"
Public Const SH_SUPPLIERS     As String = "Suppliers"
Public Const SH_HELPERS       As String = "Helpers"
Public Const SH_LOG           As String = "Log"
Public Const SH_COMPS         As String = "Comps"
Public Const SH_USERS         As String = "Users"
Public Const SH_INV           As String = "Inv"
Public Const SH_BOMS          As String = "BOMS"
Public Const SH_WOS           As String = "WOS"
Public Const SH_WOCOMPS       As String = "WOComps"
Public Const SH_DEMAND        As String = "Demand"
Public Const SH_POS           As String = "POS"
Public Const SH_POLINES       As String = "POLines"
Public Const SH_BOM_PDM001    As String = "BOM_PDM-001"
Public Const SH_BOM_TEMPLATE  As String = "BOM_TEMPLATE"

'===============================================================================
' Table (ListObject) Names  - must match ListObject.Name in each sheet
'===============================================================================

Public Const TBL_NA           As String = "TBL_NA"          ' Landing header (optional)
Public Const TBL_SCHEMA       As String = "TBL_SCHEMA"
Public Const TBL_AUTO         As String = "TBL_AUTO"
Public Const TBL_SUPPLIERS    As String = "TBL_SUPPLIERS"
Public Const TBL_HELPERS      As String = "TBL_HELPERS"
Public Const TBL_LOG          As String = "TBL_LOG"
Public Const TBL_COMPS        As String = "TBL_COMPS"
Public Const TBL_USERS        As String = "TBL_USERS"
Public Const TBL_INV          As String = "TBL_INV"
Public Const TBL_BOMS         As String = "TBL_BOMS"
Public Const TBL_WOS          As String = "TBL_WOS"
Public Const TBL_WOCOMPS      As String = "TBL_WOCOMPS"
Public Const TBL_DEMAND       As String = "TBL_DEMAND"
Public Const TBL_POS          As String = "TBL_POS"
Public Const TBL_POLINES      As String = "TBL_POLINES"
Public Const TBL_BOM_PDM001   As String = "TBL_BOM_PDM001"
Public Const TBL_BOM_TEMPLATE As String = "TBL_BOM_TEMPLATE"

'===============================================================================
' Common Column Names - Keys / IDs
'===============================================================================

Public Const COL_SUPPLIER_ID      As String = "SupplierID"
Public Const COL_SUPPLIER_NAME    As String = "SupplierName"

Public Const COL_COMP_ID          As String = "CompID"
Public Const COL_PN               As String = "OurPN"
Public Const COL_REV              As String = "OurRev"

Public Const COL_USER_ID          As String = "UserID"
Public Const COL_USER_NAME        As String = "UserName"

Public Const COL_TRANSACTION_ID   As String = "TransactionID"

Public Const COL_BOM_ID           As String = "BOMID"
Public Const COL_BOM_STATUS       As String = "BOMStatus"
Public Const COL_ASSEMBLY_ID      As String = "AssemblyID"
Public Const COL_TAID             As String = "TAID"

Public Const COL_BUILD_ID         As String = "BuildID"
Public Const COL_BUILD_STATUS     As String = "BuildStatus"

Public Const COL_PO_ID            As String = "POID"
Public Const COL_PO_NUMBER        As String = "PONumber"
Public Const COL_PO_STATUS        As String = "POStatus"
Public Const COL_PO_LINE          As String = "POLine"

'===============================================================================
' Common Column Names - Audit Fields
'===============================================================================

Public Const COL_CREATED_AT       As String = "CreatedAt"
Public Const COL_CREATED_BY       As String = "CreatedBy"
Public Const COL_UPDATED_AT       As String = "UpdatedAt"
Public Const COL_UPDATED_BY       As String = "UpdatedBy"

'===============================================================================
' Common Column Names - Log Table (TBL_LOG)
'===============================================================================

Public Const COL_LOG_TIMESTAMP    As String = "Timestamp"
Public Const COL_LOG_LEVEL        As String = "Level"
Public Const COL_LOG_PROC         As String = "Proc"
Public Const COL_LOG_MESSAGE      As String = "Message"
Public Const COL_LOG_DETAILS      As String = "Details"
Public Const COL_LOG_ERRNUM       As String = "ErrNum"
Public Const COL_LOG_ACTIVITY_ID  As String = "ActivityId"
Public Const COL_LOG_REPEAT_COUNT As String = "RepeatCount"
Public Const COL_LOG_USER_ID      As String = "UserID"
Public Const COL_LOG_WORKBOOK     As String = "Workbook"
Public Const COL_LOG_VERSION      As String = "Version"
Public Const COL_LOG_OTHER        As String = "Other"

'===============================================================================
' Common Column Names - Demand / Inventory / Quantities
'===============================================================================

Public Const COL_DESCRIPTION      As String = "Description"
Public Const COL_UOM              As String = "UOM"

Public Const COL_QOH_BEFORE       As String = "QOH_BEFORE"
Public Const COL_QOH_DELTA        As String = "ADD/SUBTRACT"
Public Const COL_QOH              As String = "QOH"
Public Const COL_ALLOCATED        As String = "ALLOCATED"
Public Const COL_NET_AVAILABLE    As String = "NetAvailable"

Public Const COL_TOTAL_DEMAND     As String = "TotalDemand"
Public Const COL_TOTAL_ONHAND     As String = "TotalOnHand"
Public Const COL_NET_DEMAND       As String = "NetDemand"

Public Const COL_QTY_PER          As String = "QtyPer"
Public Const COL_BUILD_QTY_DEMAND As String = "BuildQuantityDemand"
Public Const COL_PO_QTY           As String = "POQuantity"
Public Const COL_PRICE_PER_UOM    As String = "PricePerUOM"
Public Const COL_PO_LINE_TOTAL    As String = "POLineTotal"

'===============================================================================
' BOM lifecycle statuses
'===============================================================================

Public Const BOM_STATUS_DRAFT     As String = "DRAFT"
Public Const BOM_STATUS_LOCK      As String = "LOCK"
Public Const BOM_STATUS_OBSOLETE  As String = "OBSOLETE"

'===============================================================================
' Log Levels
'===============================================================================

Public Const LOG_LEVEL_INFO       As String = "INFO"
Public Const LOG_LEVEL_WARN       As String = "WARN"
Public Const LOG_LEVEL_ERROR      As String = "ERROR"

'===============================================================================
' Shared Messages
'===============================================================================

Public Const MSG_SCHEMA_INVALID   As String = _
    "Workbook structure does not match required schema. See Schema_Check."

Public Const MSG_CONFIRM_PROCEED  As String = _
    "This action will modify data in one or more tables. Do you want to continue?"

Public Const MSG_INTERNAL_ERROR   As String = _
    "An unexpected error occurred. See the Log sheet for details."


