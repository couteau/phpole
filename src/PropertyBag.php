<?php
namespace Cryptodira\PhpMsOle;

use function Cryptodira\PhpMsOle\Util\ReadCodePageString;

require_once('Util.php');

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


const FMTID_DocSummaryInformation   = 'd5cdd5022e9c101b939708002b2cf9ae';
const FMTID_UserDefinedProperties   = 'd5cdd5052e9c101b939708002b2cf9ae';
const FMTID_SummaryInformation      = 'f29f85e04ff91068ab9108002b27b3d9';
const FMTID_GlobalInfo              = '56616f00c15411ce855300aa00a1f95b';
const FMTID_ImageContents           = '56616400c15411ce855300aa00a1f95b';
const FMTID_ImageInfo               = '56616500c15411ce855300aa00a1f95b';


/**
 * Class for reading/manipulating property bag streams in an OLE file
 *
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
class PropertyBag
{
    /**
     * Header of the preoperty bag
     *
     * @var array
     */
    private $header;

    /**
     * The property name => value dictionary for all properties in the stream
     * @var array
     */
    private $properties;

    /**
     * The codepage in which string property values are encoded
     * @var unknown
     */
    private $codepage;

    // Unpack() format string for the property stream header
    const PropertyStreamHeaderFormat = 'v1ByteOrder/v1Version/V1SystemIdentifier/' .
        'H32CLSID/V1NumPropertySets/' .
        'V1FMTID01/v1FMTID02/v1FMTID03/H4FMTID04/H12FMTID05/' .
        'V1Offset0';

    // Unpack() format string for a single property set within the property stream
    const PropertySetHeaderFormat    = 'V1Size/V1NumProperties';

    // Unpack() format string for the individual property header
    const PropertyFormat             = 'V1Identifier/V1Offset';

    // Unpack() format string for the property value
    const PropertyValueFormat        = 'v1Type/v1Padding';

    // Unpack() format strings for individual property types
    private static $ValueFormats = [
        VT_EMPTY => '',
        VT_NULL => '',
        VT_I2 => 'C1Value1/c1Value2/v1Padding',
        VT_I4 => 'C3Value/c1Value4',
        VT_R4 => 'g1Value',
        VT_R8 => 'e1Vlaue',
        VT_CY => 'C7Value/c1Value8',
        VT_DATE => 'e1Value',
        VT_BSTR => '',
        VT_ERROR => 'V1Value',
        VT_BOOL => 'C1Value/C3',
        VT_DECIMAL => 'v1Reserved/c1Scale/c1Sign/V1Hi32/P1Lo64',
        VT_I1 => 'c1Value/C3',
        VT_UI1 => 'C1Value/C3Padding',
        VT_UI2 => 'v1Value/v1Padding',
        VT_UI4 => 'V1Value',
        VT_I8 => 'C7Value/c1Value8',
        VT_UI8 => 'P1Value',
        VT_INT => 'C3Value/c1Value4',
        VT_UINT => 'V1Value',
        VT_LPSTR => '',
        VT_LPWSTR => 'V1Size',
        VT_FILETIME => 'V1dwLowDateTime/V1dwHighDateTime',
        //VT_BLOB => '',
        //VT_STREAM => '',
        //VT_STORAGE => '',
        //VT_STREAMED_Object => '',
        //VT_STORED_Object => '',
        //VT_BLOB_Object => '',
        //VT_CF => '',
        VT_CLSID => 'H32Value',
        //VT_VERSIONED_STREAM => '',
        //VT_VECTOR => '',
        //VT_ARRAY => '',
    ];

    // Dictionary for the standard Summary Information property set
    private static $SummaryInformationDictionary = [
        0x00000001 => "CodePage",
        PIDSI_TITLE => 'Title',
        PIDSI_SUBJECT => 'Subject',
        PIDSI_AUTHOR => 'Author',
        PIDSI_KEYWORDS => 'Keywords',
        PIDSI_COMMENTS => 'Comments',
        PIDSI_TEMPLATE => 'Template',
        PIDSI_LASTAUTHOR => 'Last Author',
        PIDSI_REVNUMBER => 'Revision Number',
        PIDSI_APPNAME => 'Application Name',
        PIDSI_EDITTIME => 'Edit Time',
        PIDSI_LASTPRINTED => 'Last Printed',
        PIDSI_CREATE_DTM => 'Create Date/Time',
        PIDSI_LASTSAVE_DTM => 'Last Save Date/Time',
        PIDSI_PAGECOUNT => 'Page Count',
        PIDSI_WORDCOUNT => 'Word Count',
        PIDSI_CHARCOUNT => 'Character Count',
        PIDSI_DOC_SECURITY => 'Document Security',];

    // Dictionary for the standard Document Summary Information property set
    private static $DocumentSummaryDictionary = [
        GKPIDDSI_CODEPAGE => 'CodePage',
        GKPIDDSI_CATEGORY => 'Category',
        GKPIDDSI_PRESFORMAT => 'Presentation Format',
        GKPIDDSI_BYTECOUNT => 'Byte Count',
        GKPIDDSI_LINECOUNT => 'Line Count',
        GKPIDDSI_PARACOUNT => 'Paragraph Count',
        GKPIDDSI_SLIDECOUNT => 'Slide Count',
        GKPIDDSI_NOTECOUNT => 'Note Count',
        GKPIDDSI_HIDDENCOUNT => 'Hidden Slide Count',
        GKPIDDSI_MMCLIPCOUNT => 'Multimedia Clip Count',
        GKPIDDSI_SCALE => 'Scale',
        GKPIDDSI_HEADINGPAIR => 'Heading Pair',
        GKPIDDSI_DOCPARTS => 'Document Parts',
        GKPIDDSI_MANAGER => 'Manager',
        GKPIDDSI_COMPANY => 'Company',
        GKPIDDSI_LINKSDIRTY => 'Links Dirty',
        GKPIDDSI_CCHWITHSPACES => 'Character Count with Spaces',
        GKPIDDSI_SHAREDDOC => 'Shared Document',
        GKPIDDSI_LINKBASE => 'Link Base',
        GKPIDDSI_HLINKS => 'Hyperlinks',
        GKPIDDSI_HYPERLINKSCHANGED => 'Hyperlinks Changed',
        GKPIDDSI_VERSION => 'Application Version',
        GKPIDDSI_DIGSIG => 'Digital Signature',
        GKPIDDSI_CONTENTTYPE => 'Content Type',
        GKPIDDSI_CONTENTSTATUS => 'Content Status',
        GKPIDDSI_LANGUAGE => 'Language',
        GKPIDDSI_DOCVERSION => 'Document Version',
    ];

    /**
     * Create a new property bag and read the property sets/properties from the passed binary data
     *
     * @param string $data
     */
    public function __construct($data)
    {
        $this->header = unpack(self::PropertyStreamHeaderFormat, $data);
        $uuid = array_values(array_splice($this->header, 5, 5));
        $this->header['FMTID0'] = dechex($uuid[0]) . dechex($uuid[1]) . dechex($uuid[2]) . $uuid[3] . $uuid[4];
        if ($this->header['NumPropertySets'] == 2)
            $this->header = array_merge($this->header, unpack('H32FMTID1/V1Offset1', $data, 48));
        $this->properties = array();
        for ($i=0; $i < $this->header['NumPropertySets']; $i++) {
            $this->codepage = 1252;
            if ($this->header['FMTID' . $i ] == FMTID_SummaryInformation)
                $dictionary = self::$SummaryInformationDictionary;
            elseif ($this->header['FMTID' . $i ] == FMTID_DocSummaryInformation)
                $dictionary = self::$DocumentSummaryDictionary;
            else
                $dictionary = null;

            $offset = $this->header['Offset' . $i];
            $header = unpack(self::PropertySetHeaderFormat, $data, $offset);
            $props = array();
            for ($j=0; $j < $header['NumProperties']; $j++) {
                [, $id, $propoffset] = unpack('V2', $data, $offset + 8 + $j*8);
                if ($id == PID_DICTIONARY) {
                    $dictionary = $this->ReadDictionary($data, $offset + $propoffset);
                    $value = null;
                }
                else {
                    $propoffset += $offset;
                    $value = $this->GetPropertyValue($data, $propoffset);
                }

                $props[$id] = $value;

                if ($id == PID_CODEPAGE)
                    $this->codepage = $value;
            }

            // if the property set contained a dictionary, substitute property names for ids
            if ($dictionary) {
                foreach ($props as $key => $value) {
                    if (isset($dictionary[$key])) {
                        $props[$dictionary[$key]] = $value;
                        unset($props[$key]);
                    }
                }
            }

            $this->properties = array_merge($this->properties, $props);
        }
    }

    /**
     * Convert the encoding of the passed string from the codepage for this property set to UTF-8
     *
     * @param string $str
     * @return string
     */
    private function convert_encoding($str)
    {
        if (isset(mb_encodings[$this->codepage]))
            $out = mb_convert_encoding($str, 'UTF-8', mb_encodings[$this->codepage]);
        elseif (isset(iconv_encodings[$this->codepage])) {
            $out = iconv(iconv_encodings[$this->codepage], 'UTF-8//TRANSLIT', $str);
            if (!$out)
                $out = iconv(iconv_encodings[$this->codepage], 'UTF-8//IGNORE', $str);
        }
        elseif ($this->codepage >= 437 && $this->codepage <= 1258) {
            $cp = 'CP' . $this->codepage;
            $out = iconv($cp, 'UTF-8//TRANSLIT', $str);
            if (!$out)
                $out = iconv($cp, 'UTF-8//IGNORE', $str);
            if (!$out)
                $out = $str;
        }
        else
            $out = $str;

        return rtrim($out, "\0");
    }

    /**
     * Read a dictonary from a dictionary property
     *
     * @param string $data
     * @param int $offset
     * @return string[]
     */
    private function ReadDictionary($data, $offset)
    {
        $dictionary = array();
        $numentries = unpack('V1', $data, $offset)[1];
        for ($i = 0, $o = $offset + 4; $i < $numentries; $i++) {
            [, $id, $length] = unpack('V2', $data, $o);
            if ($this->codepage == "UTF-16LE")
                $length = 2 * $length;
            $name = mb_convert_encoding(substr($data, $o+8, $length), 'UTF-8', $this->codpage);
            $dictionary[$id] = $name;
            $o += 8 + $length + ((4 - ($value['Size'] % 4) % 4));
        }

        return $dictionary;
    }

    /**
     * Read a property value from the stream starting at offset, and update $offset to end of value data
     *
     * @param string $data
     * @param int $offset
     * @param int $type
     * @return mixed
     */
    private function GetPropertyValue($data, &$offset, $type = null)
    {
        if (!$type || $type == VT_VARIANT) {
            $type = unpack(self::PropertyValueFormat, $data, $offset)['Type'];
            $offset += 4;
        }


        if (isset(self::$ValueFormats[$type]) && !empty(self::$ValueFormats[$type]))
            $value = unpack(self::$ValueFormats[$type], $data, $offset);
        else
            $value = null;

        switch ($type) {
            case VT_EMPTY:
            case VT_NULL:
                $value= null;
                break;
            case VT_I2:
                $value = ($value['Value2'] << 8) | $value['Value1'];
                $offset += 4;
                break;
            case VT_I4:
            case VT_INT:
                $value = ($value['Value4'] << 24) | ($value['Value3'] << 16) | ($value['Value2'] << 8) | $value['Value1'];
                $offset += 4;
                break;
            case VT_DATE:
                $value = MSDATETIME_BASE + $value['Value'] * 86400;
                $offset += 8;
                break;
            case VT_BOOL:
                $value = $value['Value'] ? true : false;
                $offset += 4;
                break;
            case VT_DECIMAL:
                $v = bcdiv(bcadd(bcmul($value['Hi32'], bcpow(2, 64)), $value['Lo64']), bcpow(10,$value['Scale']), $value['Scale']);
                if ($value['Sign'] == 0x80)
                    $value = bcmul(-1, $v, $value['Scale']);
                else
                    $value = $v;
                $offset += 16;
                break;
            case VT_CY:
            case VT_I8:
                $value = ($value['Value8'] << 56) | ($value['Value7'] << 48) | ($value['Value6'] << 40) | ($value['Value5'] << 32) |
                ($value['Value4'] << 24) | ($value['Value3'] << 16) | ($value['Value2'] << 8) | $value['Value1'];
                if ($type == VT_CY)
                    $value /= 10000;
                $offset += 8;
                break;
            case VT_BSTR:
            case VT_LPSTR:
                $bytesread = 0;
                $value = ReadCodePageString($data, $offset, $bytesread, $this->codepage);
                //$offset += $bytesread + ((4 - ($value['Size'] % 4) % 4));
                $offset += $bytesread; //Padding to 4 byte boundary doesn't seem to be happening notwithstanding the spec
                break;
            case VT_LPWSTR:
                $bytesread = 0;
                $value = ReadUnicodeString($data, $offset, $bytesread);
                $offset += $bytesread;
                break;
            case VT_FILETIME:
                $ft =  ($value['dwHighDateTime'] << 32) | $value['dwLowDateTime'];
                if ($ft==0)
                    $value = null;
                else
                    $value = FILETIME_BASE + $ft / 10000000;
                break;
            // Unsupported Types
            case VT_BLOB:
            case VT_STREAM:
            case VT_STORAGE:
            case VT_STREAMED_Object:
            case VT_STORED_Object:
            case VT_BLOB_Object:
            case VT_CF:
            case VT_VERSIONED_STREAM:
                $value = null;
                break;
            default:
                if ($type & VT_VECTOR) {
                    $size = unpack('V1', $data, $offset)[1];
                    $value = array();
                    $offset += 4;
                    for ($i = 0; $i < $size; $i++)
                        $value[] = $this->GetPropertyValue($data, $offset, $type & 0xFFF);
                }
                else if ($type & VT_ARRAY)
                    $value = null; // Not supporting Arrays for now
                else
                    $value = $value['Value'];
        }

        return $value;
    }
}

