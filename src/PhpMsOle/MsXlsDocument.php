<?php
namespace Cryptodira\PhpMsOle;

use function Cryptodira\PhpMsOle\Util\ReadCodePageString;
use function Cryptodira\PhpMsOle\Util\ConvertEncoding;
use \phpseclib\Crypt\RC4;


include_once('Util.php');

/**
 *
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
class MsXlsDocument
{
    private $OLEDocument;
    private $streamid;
    private $summarystreamid;
    private $docinfostreamid;
    private $summaryinfo;
    private $docsummaryinfo;
    private $stream;
    private $BOF;
    private $worksheets;

    private $encryptiontype;
    private $encryptionkey;
    private $verificationbytes;

    const XLS_BIFF8                     = 0x0600;
    const XLS_BIFF7                     = 0x0500;

    const ENC_BLOCK_SIZE = 1024;

    const RT_AlRuns = 4176;
    const RT_Area = 4122;
    const RT_AreaFormat = 4106;
    const RT_Array = 545;
    const RT_AttachedLabel = 4108;
    const RT_AutoFilter = 158;
    const RT_AutoFilter12 = 2174;
    const RT_AutoFilterInfo = 157;
    const RT_AxcExt = 4194;
    const RT_AxesUsed = 4166;
    const RT_Axis = 4125;
    const RT_AxisLine = 4129;
    const RT_AxisParent = 4161;
    const RT_Backup = 64;
    const RT_Bar = 4119;
    const RT_BCUsrs = 407;
    const RT_Begin = 4147;
    const RT_BigName = 1048;
    const RT_BkHim = 233;
    const RT_Blank = 513;
    const RT_BOF = 2057;
    const RT_BookBool = 218;
    const RT_BookExt = 2147;
    const RT_BoolErr = 517;
    const RT_BopPop = 4193;
    const RT_BopPopCustom = 4199;
    const RT_BottomMargin = 41;
    const RT_BoundSheet8 = 133;
    const RT_BRAI = 4177;
    const RT_BuiltInFnGroupCount = 156;
    const RT_CalcCount = 12;
    const RT_CalcDelta = 16;
    const RT_CalcIter = 17;
    const RT_CalcMode = 13;
    const RT_CalcPrecision = 14;
    const RT_CalcRefMode = 15;
    const RT_CalcSaveRecalc = 95;
    const RT_CatLab = 2134;
    const RT_CatSerRange = 4128;
    const RT_CbUsr = 402;
    const RT_CellWatch = 2156;
    const RT_CF = 433;
    const RT_CF12 = 2170;
    const RT_CFEx = 2171;
    const RT_Chart = 4098;
    const RT_Chart3d = 4154;
    const RT_Chart3DBarShape = 4191;
    const RT_ChartFormat = 4116;
    const RT_ChartFrtInfo = 2128;
    const RT_ClrtClient = 4188;
    const RT_CodeName = 442;
    const RT_CodePage = 66;
    const RT_ColInfo = 125;
    const RT_Compat12 = 2188;
    const RT_CompressPictures = 2203;
    const RT_CondFmt = 432;
    const RT_CondFmt12 = 2169;
    const RT_Continue = 60;
    const RT_ContinueBigName = 1084;
    const RT_ContinueFrt = 2066;
    const RT_ContinueFrt11 = 2165;
    const RT_ContinueFrt12 = 2175;
    const RT_Country = 140;
    const RT_CrErr = 2149;
    const RT_CRN = 90;
    const RT_CrtLayout12 = 2205;
    const RT_CrtLayout12A = 2215;
    const RT_CrtLine = 4124;
    const RT_CrtLink = 4130;
    const RT_CrtMlFrt = 2206;
    const RT_CrtMlFrtContinue = 2207;
    const RT_CUsr = 401;
    const RT_Dat = 4195;
    const RT_DataFormat = 4102;
    const RT_DataLabExt = 2154;
    const RT_DataLabExtContents = 2155;
    const RT_Date1904 = 34;
    const RT_DBCell = 215;
    const RT_DbOrParamQry = 220;
    const RT_DBQueryExt = 2051;
    const RT_DCon = 80;
    const RT_DconBin = 437;
    const RT_DConn = 2166;
    const RT_DConName = 82;
    const RT_DConRef = 81;
    const RT_DefaultRowHeight = 549;
    const RT_DefaultText = 4132;
    const RT_DefColWidth = 85;
    const RT_Dimensions = 512;
    const RT_DocRoute = 184;
    const RT_DropBar = 4157;
    const RT_DropDownObjIds = 2164;
    const RT_DSF = 353;
    const RT_Dv = 446;
    const RT_DVal = 434;
    const RT_DXF = 2189;
    const RT_DxGCol = 153;
    const RT_End = 4148;
    const RT_EndBlock = 2131;
    const RT_EndObject = 2133;
    const RT_EntExU2 = 450;
    const RT_EOF = 10;
    const RT_Excel9File = 448;
    const RT_ExternName = 35;
    const RT_ExternSheet = 23;
    const RT_ExtSST = 255;
    const RT_ExtString = 2052;
    const RT_Fbi = 4192;
    const RT_Fbi2 = 4200;
    const RT_Feat = 2152;
    const RT_FeatHdr = 2151;
    const RT_FeatHdr11 = 2161;
    const RT_Feature11 = 2162;
    const RT_Feature12 = 2168;
    const RT_FileLock = 405;
    const RT_FilePass = 47;
    const RT_FileSharing = 91;
    const RT_FilterMode = 155;
    const RT_FnGroupName = 154;
    const RT_FnGrp12 = 2200;
    const RT_Font = 49;
    const RT_FontX = 4134;
    const RT_Footer = 21;
    const RT_ForceFullCalculation = 2211;
    const RT_Format = 1054;
    const RT_Formula = 6;
    const RT_Frame = 4146;
    const RT_FrtFontList = 2138;
    const RT_FrtWrapper = 2129;
    const RT_GelFrame = 4198;
    const RT_GridSet = 130;
    const RT_GUIDTypeLib = 2199;
    const RT_Guts = 128;
    const RT_HCenter = 131;
    const RT_Header = 20;
    const RT_HeaderFooter = 2204;
    const RT_HFPicture = 2150;
    const RT_HideObj = 141;
    const RT_HLink = 440;
    const RT_HLinkTooltip = 2048;
    const RT_HorizontalPageBreaks = 27;
    const RT_IFmtRecord = 4174;
    const RT_Index = 523;
    const RT_InterfaceEnd = 226;
    const RT_InterfaceHdr = 225;
    const RT_Intl = 97;
    const RT_Label = 516;
    const RT_LabelSst = 253;
    const RT_Lbl = 24;
    const RT_LeftMargin = 38;
    const RT_Legend = 4117;
    const RT_LegendException = 4163;
    const RT_Lel = 441;
    const RT_Line = 4120;
    const RT_LineFormat = 4103;
    const RT_List12 = 2167;
    const RT_LPr = 152;
    const RT_LRng = 351;
    const RT_MarkerFormat = 4105;
    const RT_MDB = 2186;
    const RT_MDTInfo = 2180;
    const RT_MDXKPI = 2185;
    const RT_MDXProp = 2184;
    const RT_MDXSet = 2183;
    const RT_MDXStr = 2181;
    const RT_MDXTuple = 2182;
    const RT_MergeCells = 229;
    const RT_Mms = 193;
    const RT_MsoDrawing = 236;
    const RT_MsoDrawingGroup = 235;
    const RT_MsoDrawingSelection = 237;
    const RT_MTRSettings = 2202;
    const RT_MulBlank = 190;
    const RT_MulRk = 189;
    const RT_NameCmt = 2196;
    const RT_NameFnGrp12 = 2201;
    const RT_NamePublish = 2195;
    const RT_Note = 28;
    const RT_Number = 515;
    const RT_Obj = 93;
    const RT_ObjectLink = 4135;
    const RT_ObjProtect = 99;
    const RT_ObNoMacros = 445;
    const RT_ObProj = 211;
    const RT_OleDbConn = 2058;
    const RT_OleObjectSize = 222;
    const RT_Palette = 146;
    const RT_Pane = 65;
    const RT_Password = 19;
    const RT_PhoneticInfo = 239;
    const RT_PicF = 4156;
    const RT_Pie = 4121;
    const RT_PieFormat = 4107;
    const RT_PivotChartBits = 2137;
    const RT_PlotArea = 4149;
    const RT_PlotGrowth = 4196;
    const RT_Pls = 77;
    const RT_PLV = 2187;
    const RT_Pos = 4175;
    const RT_PrintGrid = 43;
    const RT_PrintRowCol = 42;
    const RT_PrintSize = 51;
    const RT_Prot4Rev = 431;
    const RT_Prot4RevPass = 444;
    const RT_Protect = 18;
    const RT_Qsi = 429;
    const RT_Qsif = 2055;
    const RT_Qsir = 2054;
    const RT_QsiSXTag = 2050;
    const RT_Radar = 4158;
    const RT_RadarArea = 4160;
    const RT_RealTimeData = 2067;
    const RT_RecalcId = 449;
    const RT_RecipName = 185;
    const RT_RefreshAll = 439;
    const RT_RichTextStream = 2214;
    const RT_RightMargin = 39;
    const RT_RK = 638;
    const RT_Row = 520;
    const RT_RRAutoFmt = 331;
    const RT_RRDChgCell = 315;
    const RT_RRDConflict = 338;
    const RT_RRDDefName = 339;
    const RT_RRDHead = 312;
    const RT_RRDInfo = 406;
    const RT_RRDInsDel = 311;
    const RT_RRDInsDelBegin = 336;
    const RT_RRDInsDelEnd = 337;
    const RT_RRDMove = 320;
    const RT_RRDMoveBegin = 334;
    const RT_RRDMoveEnd = 335;
    const RT_RRDRenSheet = 318;
    const RT_RRDRstEtxp = 340;
    const RT_RRDTQSIF = 2056;
    const RT_RRDUserView = 428;
    const RT_RRFormat = 330;
    const RT_RRInsertSh = 333;
    const RT_RRSort = 319;
    const RT_RRTabId = 317;
    const RT_SBaseRef = 4168;
    const RT_Scatter = 4123;
    const RT_SCENARIO = 175;
    const RT_ScenarioProtect = 221;
    const RT_ScenMan = 174;
    const RT_Scl = 160;
    const RT_Selection = 29;
    const RT_SerAuxErrBar = 4187;
    const RT_SerAuxTrend = 4171;
    const RT_SerFmt = 4189;
    const RT_Series = 4099;
    const RT_SeriesList = 4118;
    const RT_SeriesText = 4109;
    const RT_SerParent = 4170;
    const RT_SerToCrt = 4165;
    const RT_Setup = 161;
    const RT_ShapePropsStream = 2212;
    const RT_SheetExt = 2146;
    const RT_ShrFmla = 1212;
    const RT_ShtProps = 4164;
    const RT_SIIndex = 4197;
    const RT_Sort = 144;
    const RT_SortData = 2197;
    const RT_SST = 252;
    const RT_StartBlock = 2130;
    const RT_StartObject = 2132;
    const RT_String = 519;
    const RT_Style = 659;
    const RT_StyleExt = 2194;
    const RT_SupBook = 430;
    const RT_Surf = 4159;
    const RT_SXAddl = 2148;
    const RT_SxBool = 202;
    const RT_SXDB = 198;
    const RT_SXDBB = 200;
    const RT_SXDBEx = 290;
    const RT_SXDI = 197;
    const RT_SXDtr = 206;
    const RT_SxDXF = 244;
    const RT_SxErr = 203;
    const RT_SXEx = 241;
    const RT_SXFDB = 199;
    const RT_SXFDBType = 443;
    const RT_SxFilt = 242;
    const RT_SxFmla = 249;
    const RT_SxFormat = 251;
    const RT_SXFormula = 259;
    const RT_SXInt = 204;
    const RT_SxIsxoper = 217;
    const RT_SxItm = 245;
    const RT_SxIvd = 180;
    const RT_SXLI = 181;
    const RT_SxName = 246;
    const RT_SxNil = 207;
    const RT_SXNum = 201;
    const RT_SXPair = 248;
    const RT_SXPI = 182;
    const RT_SXPIEx = 2062;
    const RT_SXRng = 216;
    const RT_SxRule = 240;
    const RT_SxSelect = 247;
    const RT_SXStreamID = 213;
    const RT_SXString = 205;
    const RT_SXTbl = 208;
    const RT_SxTbpg = 210;
    const RT_SXTBRGIITM = 209;
    const RT_SXTH = 2061;
    const RT_Sxvd = 177;
    const RT_SXVDEx = 256;
    const RT_SXVDTEx = 2063;
    const RT_SXVI = 178;
    const RT_SxView = 176;
    const RT_SXViewEx = 2060;
    const RT_SXViewEx9 = 2064;
    const RT_SXViewLink = 2136;
    const RT_SXVS = 227;
    const RT_Sync = 151;
    const RT_Table = 566;
    const RT_TableStyle = 2191;
    const RT_TableStyleElement = 2192;
    const RT_TableStyles = 2190;
    const RT_Template = 96;
    const RT_Text = 4133;
    const RT_TextPropsStream = 2213;
    const RT_Theme = 2198;
    const RT_Tick = 4126;
    const RT_TopMargin = 40;
    const RT_TxO = 438;
    const RT_TxtQry = 2053;
    const RT_Uncalced = 94;
    const RT_Units = 4097;
    const RT_UserBView = 425;
    const RT_UserSViewBegin = 426;
    const RT_UserSViewBegin_Chart = 426;
    const RT_UserSViewEnd = 427;
    const RT_UsesELFs = 352;
    const RT_UsrChk = 408;
    const RT_UsrExcl = 404;
    const RT_UsrInfo = 403;
    const RT_ValueRange = 4127;
    const RT_VCenter = 132;
    const RT_VerticalPageBreaks = 26;
    const RT_WebPub = 2049;
    const RT_Window1 = 61;
    const RT_Window2 = 574;
    const RT_WinProtect = 25;
    const RT_WOpt = 2059;
    const RT_WriteAccess = 92;
    const RT_WriteProtect = 134;
    const RT_WsBool = 129;
    const RT_XCT = 89;
    const RT_XF = 224;
    const RT_XFCRC = 2172;
    const RT_XFExt = 2173;
    const RT_YMult = 2135;

    const ST_GLOBAL = 0x0005;
    const ST_SHEET = 0x0010;
    const ST_CHART = 0x0020;
    const ST_MACRO = 0x0040;

    const BOFFormat =
    'v1vers/' .
    'v1dt/' .
    'v1rupBuild/' .
    'v1rupYear/' .
    'V2flags';

    const RecordReaders = array(
        self::RT_BoundSheet8 => 'ReadBoundSheet8',
        self::RT_BOF => 'ReadBOF',
        self::RT_FilePass => 'ReadFilePass',
        self::RT_RRTabId => 'ReadRRTabId',
    );

    /**
     */
    public function __construct($file)
    {
        $this->OLEDocument = new OLEDocument();
        if (is_string($file))
            $this->OLEDocument->CreateFromFile($file);
        elseif (is_resource($file) && get_resource_type($file) == 'stream')
            $this->OLEDocument->CreateFromStream($file);
        else
            throw new \Exception("Must pass a file name or stream to Word97Document constructor");

        $this->worksheets = array();
        $this->streamid = $this->OLEDocument->GetDocumentStream();
        $this->summarystreamid = $this->OLEDocument->FindStreamByName(chr(5) . 'SummaryInformation');
        $this->docinfostreamid = $this->OLEDocument->FindStreamByName(chr(5) . 'DocumentSummaryInformation');
        $this->docsummaryinfo = null;
        $this->summaryinfo = null;
        $this->encryptiontype = -1;

        $this->stream = new OLEStreamReader($this->OLEDocument, $this->streamid);

        // Read the global substream
        $this->ReadSubstream();
    }

    public function GetDocumentSummary()
    {
        if (is_null($this->docsummaryinfo))
            $this->docsummaryinfo = new PropertyBag($this->OLEDocument->Read($this->docinfostreamid));
        return $this->docsummaryinfo;
    }

    public function GetSummaryInfo()
    {
        if (is_null($this->summaryinfo))
            $this->summaryinfo = new PropertyBag($this->OLEDocument->Read($this->summarystreamid));
        return $this->summaryinfo;
    }

    private function ReadData($size)
    {
        $data = $this->stream->Read($size);

        if ($this->encryptiontype == 0x0001) {
            $out = '';
            $pos = $this->stream->Tell();
            $oldblock = intdiv($this->encpos, self::ENC_BLOCK_SIZE);
            $block = intdiv($pos, self::ENC_BLOCK_SIZE);
            $endblock = intdiv($pos + $size, self::ENC_BLOCK_SIZE);

            // Spin an RC4 decryptor to the right spot. If we have a decryptor sitting
            // at a point earlier in the current block, re-use it as we can save some time.
            if ($block != $oldblock || $pos < $this->encpos) {
                $this->RekeyRC4($block);
                $step = $pos % self::ENC_BLOCK_SIZE;
            } else {
                $step = $pos - $this->encpos;
            }
            $this->rc4->decrypt(str_repeat("\0", $step));

            // Decrypt record data (re-keying at the end of every block)
            while ($block != $endblock) {
                $step = self::ENC_BLOCK_SIZE - ($pos % self::ENC_BLOCK_SIZE);
                $out .= $this->rc4->decrypt(substr($data, 0, $step));
                $data = substr($data, $step);
                $pos += $step;
                $size -= $step;
                ++$block;
                $this->RekeyRC4($block);
            }
            $out .= $this->rc4Key->RC4(substr($data, 0, $size));
            $this->encpos = $pos + $size;
            $data = $out;
        }
        elseif ($this->encryptiontype == 0x0002 || $this->encryptiontype == 0x0003 || $this->encryptiontype == 0x0004) {
            throw new \Exception('CryptoAPI encryption not supported');
        }
        elseif ($this->encryptiontype == 0x0000) { // XOR obfuscation
            throw new \Exception('XOR encryption not supported');
        }

        return $data;
    }

    private function ReadShortXLUnicodeString($buffer, $offset, &$bytesread = null)
    {
        global $CODEPAGE;

        [,$size,$format] = unpack('C2', $buffer, $offset);
        if ($format & 0x80) {
            $str = mb_convert_encoding(substr($buffer, 2, $size*2), 'UTF-8', 'UTF-16LE');
            $bytesread = 2 + $size * 2;
        }
        else {
            $str = ConvertEncoding(substr($buffer, 2, $size));
            $bytesread = 2 + $size;
        }

        return $str;
    }

    private function ReadBOF($type, $size)
    {
        $data = $this->stream->Read($size);
        $BOF = unpack(self::BOFFormat, $data);

        switch ($BOF['dt']) {
            case  self::ST_GLOBAL:
                if (($BOF['vers'] != self::XLS_BIFF8) && ($BOF['vers'] != self::XLS_BIFF7)) {
                    throw new \Exception('Cannot read this Excel file. Version is too old.');
                }
                $this->BOF = $BOF;
                unset($this->sheetids);
                unset($this->sheetindex);
                break;
            default:
                // nop
        }
    }

    private function RekeyRC4($block)
    {
        $hash = substr($this->hash, 0, 5) . pack('V', $block);
        $this->rc4->setKey(md5($hash, TRUE));
    }

    private function VerifyKey($pw, $salt, $verifier, $verifierhash)
    {
        $password = $pw;
        $pw = mb_convert_encoding($pw . chr(0), 'UTF-16LE');

        $pwarray = str_repeat("\0", 64);

        $iMax = strlen($password);
        for ($i = 0; $i < $iMax; ++$i) {
            $o = ord(substr($password, $i, 1));
            $pwarray[2 * $i] = chr($o & 0xff);
            $pwarray[2 * $i + 1] = chr(($o >> 8) & 0xff);
        }
        $pwarray[2 * $i] = chr(0x80);
        $pwarray[56] = chr(($i << 4) & 0xff);


        $h = md5($pw, TRUE);
        $ih = substr($h, 0, 5) . $salt;
        $ib = str_repeat($ih, 16);
        $h1 = md5($ib, TRUE);
        $this->hash = $h1;

        $h = md5($pwarray, TRUE);

        $offset = 0;
        $keyoffset = 0;
        $tocopy = 5;
        $ib = '';

        while ($offset != 16) {
            if ((64 - $offset) < 5) {
                $tocopy = 64 - $offset;
            }
            for ($i = 0; $i <= $tocopy; ++$i) {
                $pwarray[$offset + $i] = $h[$keyoffset + $i];
            }
            $offset += $tocopy;

            if ($offset == 64) {
                $ib .= $pwarray;
                $keyoffset = $tocopy;
                $tocopy = 5 - $tocopy;
                $offset = 0;

                continue;
            }

            $keyoffset = 0;
            $tocopy = 5;
            for ($i = 0; $i < 16; ++$i) {
                $pwarray[$offset + $i] = $salt[$i];
            }
            $offset += 16;
        }

        $pwarray[16] = "\x80";
        for ($i = 0; $i < 47; ++$i) {
            $pwarray[17 + $i] = "\0";
        }
        $pwarray[56] = "\x80";
        $pwarray[57] = "\x0a";

        $this->hash = md5($ib . $pwarray);

        $this->rc4 = new RC4();
        $this->RekeyRC4(0);

        $v = $this->rc4->decrypt($verifier);
        $vh = $this->rc4->decrypt($verifierhash);

        return (md5($v, TRUE) === $vh);
    }

    const EncryptionHeaderFormat =
        'V1Flags/V1SizeExtra/V1AlgID/V1AlgIDHash/V1KeySize/V1ProviderType/V1Reserved1/V1Reserved2';

    private function ReadFilePass($type, $size)
    {
        $data = $this->stream->Read($size);
        $wEncryptionType = unpack('v1', $data)[1];
        switch ($wEncryptionType) {
            case 0x0000:
                //read XORObfjuscation structure
                [,$this->ecyrpytionkey, $this->verificationbytes] = unpack('v2', $data, 2);
                break;
            case 0x0001:
                //read RC4 structure
                [,$vMajor,$vMinor] = unpack('v2', $data, 2);
                switch ($vMajor) {
                    case 0x0001:
                        $salt = substr($data, 6, 16);
                        $encryptedverifier = substr($data, 22, 16);
                        $encryptedverfierhash = substr($data, 38, 16);
                        if (!$this->VerifyKey('VelvetSweatshop', $salt, $encryptedverifier, $encryptedverfierhash))
                            throw new \Exception('Unable to verify RC4 key.');
                        break;
                    case 0x0002:
                    case 0x0003:
                    case 0x0004:
                        if ($vMinor != 2)
                            throw new \Exception('Invalid RC4 encryption minor version.');
                        [,$headerflags,$headersize] = upack('V2', $data, 6);
                        $encryptionheader = unpack(self::EncryptionHeaderFormat, $data, 14);
                        $cspname = ReadShortXUnicodeString($data, 32);
                        break;
                    default:
                        throw new \Exception('Invalid RC4 encryption major version.');
                }

                break;
            default:
                throw new \Exception('Invalid encryption type.');
        }
        $this->encryptiontype = $wEncryptionType;

        return $type;
    }

    private function ReadRRTabId($type, $size)
    {
        if ($size != 0) {
            $data = $this->ReadData($size);
            $this->sheetids = array_values(unpack('v' . ($size / 2), $data));
            $this->sheetindex = 0;
        }
        return $type;
    }

    const SheetInfoFormat =
    'V1Offset/C1Visibility/C1SheetType';

    private function ReadBoundSheet8($type, $size)
    {
        $data = $this->ReadData($size);
        $ws = unpack(self::SheetInfoFormat, $data);

        $bytesread = 0;
        $ws['Sheetname'] = $this->ReadShortXLUnicodeString($data, 6, $bytesread);

        if (isset($this->sheetids))
            $this->workshets[$this->sheetids[$this->sheetindex++]] = $ws;
        else
            $this->worksheets[] = $ws;
        return $type;
    }

    private function ReadRecord()
    {
        $data = $this->stream->Read(4);
        [,$type,$size] = unpack('v2', $data);

        if (isset(self::RecordReaders[$type])) {
            call_user_func([ $this, self::RecordReaders[$type] ], $type, $size);
        }
        else {
            // we don't know how to handle this record type; skip and ignore content, if any
            if ($size != 0)
                $this->stream->Seek($size, SEEK_CUR);
        }

        return $type;
    }

    private function ReadSubstream($offset=null)
    {
        if (!is_null($offset))
            $this->stream->Seek($offset);

        while (!$this->stream->EOF() && $this->ReadRecord() != self::RT_EOF);
    }
}

