<?php
namespace Cryptodira\PhpOle;

class Ole {
    // Unix TimeStamp for January 1, 1601
    const FILETIME_BASE = -11644473600;

    // Unix TimeSamp for December 30, 1899
    const MSDATETIME_BASE = -2209161600;

    // Microsoft Variant datatype codes
    const VT_EMPTY = 0x0000;
    const VT_NULL = 0x0001;
    const VT_I2 = 0x0002;
    const VT_I4 = 0x0003;
    const VT_R4 = 0x0004;
    const VT_R8 = 0x0005;
    const VT_CY = 0x0006;
    const VT_DATE = 0x0007;
    const VT_BSTR = 0x0008;
    const VT_ERROR = 0x000A;
    const VT_BOOL = 0x000B;
    const VT_VARIANT = 0x000C;
    const VT_DECIMAL = 0x000E;
    const VT_I1 = 0x0010;
    const VT_UI1 = 0x0011;
    const VT_UI2 = 0x0012;
    const VT_UI4 = 0x0013;
    const VT_I8 = 0x0014;
    const VT_UI8 = 0x0015;
    const VT_INT = 0x0016;
    const VT_UINT = 0x0017;
    const VT_LPSTR = 0x001E;
    const VT_LPWSTR = 0x001F;
    const VT_FILETIME = 0x0040;
    const VT_BLOB = 0x0041;
    const VT_STREAM = 0x0042;
    const VT_STORAGE = 0x0043;
    const VT_STREAMED_Object = 0x0044;
    const VT_STORED_Object = 0x0045;
    const VT_BLOB_Object = 0x0046;
    const VT_CF = 0x0047;
    const VT_CLSID = 0x0048;
    const VT_VERSIONED_STREAM = 0x0049;
    const VT_BINARY = 0x0102;
    const VT_VECTOR = 0x1000;
    const VT_ARRAY = 0x2000;

    // Property IDs
    const PID_DICTIONARY = 0x00000000;
    const PID_CODEPAGE = 0x00000001;
    
    const PIDSI_TITLE = 0x00000002;
    const PIDSI_SUBJECT = 0x00000003;
    const PIDSI_AUTHOR = 0x00000004;
    const PIDSI_KEYWORDS = 0x00000005;
    const PIDSI_COMMENTS = 0x00000006;
    const PIDSI_TEMPLATE = 0x00000007;
    const PIDSI_LASTAUTHOR = 0x00000008;
    const PIDSI_REVNUMBER = 0x00000009;
    const PIDSI_APPNAME = 0x00000012;
    const PIDSI_EDITTIME = 0x0000000A;
    const PIDSI_LASTPRINTED = 0x0000000B;
    const PIDSI_CREATE_DTM = 0x0000000C;
    const PIDSI_LASTSAVE_DTM = 0x0000000D;
    const PIDSI_PAGECOUNT = 0x0000000E;
    const PIDSI_WORDCOUNT = 0x0000000F;
    const PIDSI_CHARCOUNT = 0x00000010;
    const PIDSI_DOC_SECURITY = 0x00000013;

    const GKPIDDSI_CODEPAGE = 0x00000001;
    const GKPIDDSI_CATEGORY = 0x00000002;
    const GKPIDDSI_PRESFORMAT = 0x00000003;
    const GKPIDDSI_BYTECOUNT = 0x00000004;
    const GKPIDDSI_LINECOUNT = 0x00000005;
    const GKPIDDSI_PARACOUNT = 0x00000006;
    const GKPIDDSI_SLIDECOUNT = 0x00000007;
    const GKPIDDSI_NOTECOUNT = 0x00000008;
    const GKPIDDSI_HIDDENCOUNT = 0x00000009;
    const GKPIDDSI_MMCLIPCOUNT = 0x0000000A;
    const GKPIDDSI_SCALE = 0x0000000B;
    const GKPIDDSI_HEADINGPAIR = 0x0000000C;
    const GKPIDDSI_DOCPARTS = 0x0000000D;
    const GKPIDDSI_MANAGER = 0x0000000E;
    const GKPIDDSI_COMPANY = 0x0000000F;
    const GKPIDDSI_LINKSDIRTY = 0x00000010;
    const GKPIDDSI_CCHWITHSPACES = 0x00000011;
    const GKPIDDSI_SHAREDDOC = 0x00000013;
    const GKPIDDSI_LINKBASE = 0x00000014;
    const GKPIDDSI_HLINKS = 0x00000015;
    const GKPIDDSI_HYPERLINKSCHANGED = 0x00000016;
    const GKPIDDSI_VERSION = 0x00000017;
    const GKPIDDSI_DIGSIG = 0x00000018;
    const GKPIDDSI_CONTENTTYPE = 0x0000001A;
    const GKPIDDSI_CONTENTSTATUS = 0x0000001B;
    const GKPIDDSI_LANGUAGE = 0x0000001C;
    const GKPIDDSI_DOCVERSION = 0x0000001D;

    const FMTID_DocSummaryInformation = 'd5cdd5022e9c101b939708002b2cf9ae';
    const FMTID_UserDefinedProperties = 'd5cdd5052e9c101b939708002b2cf9ae';
    const FMTID_SummaryInformation = 'f29f85e04ff91068ab9108002b27b3d9';
    const FMTID_GlobalInfo = '56616f00c15411ce855300aa00a1f95b';
    const FMTID_ImageContents = '56616400c15411ce855300aa00a1f95b';
    const FMTID_ImageInfo = '56616500c15411ce855300aa00a1f95b';
}