<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

namespace Cryptodira\PhpMsOle;

/**
 * Microsoft PowerPoint 97 document
 *
 * @author Stuart C. Naifeh
 */
class MsPptDocument
{
    const RT_ANIMATIONINFO = 0x1014;
    const RT_ANIMATIONINFOATOM = 0x0FF1;
    const RT_BINARYTAGDATABLOB = 0x138B;
    const RT_BLIPCOLLECTION9 = 0x07F8;
    const RT_BLIPENTITY9ATOM = 0x07F9;
    const RT_BOOKMARKCOLLECTION = 0x07E3;
    const RT_BOOKMARKENTITYATOM = 0x0FD0;
    const RT_BOOKMARKSEEDATOM = 0x07E9;
    const RT_BROADCASTDOCINFO9 = 0x177E;
    const RT_BROADCASTDOCINFO9ATOM = 0x177F;
    const RT_BUILDATOM = 0x2B03;
    const RT_BUILDLIST = 0x2B02;
    const RT_CHARTBUILD = 0x2B04;
    const RT_CHARTBUILDATOM = 0x2B05;
    const RT_COLORSCHEMEATOM = 0x07F0;
    const RT_COMMENT10 = 0x2EE0;
    const RT_COMMENT10ATOM = 0x2EE1;
    const RT_COMMENTINDEX10 = 0x2EE4;
    const RT_COMMENTINDEX10ATOM = 0x2EE5;
    const RT_CRYPTSESSION10CONTAINER = 0x2F14;
    const RT_CURRENTUSERATOM = 0x0FF6;
    const RT_CSTRING = 0x0FBA;
    const RT_DATETIMEMETACHARATOM = 0x0FF7;
    const RT_DEFAULTRULERATOM = 0x0FAB;
    const RT_DOCROUTINGSLIPATOM = 0x0406;
    const RT_DIAGRAMBUILD = 0x2B06;
    const RT_DIAGRAMBUILDATOM = 0x2B07;
    const RT_DIFF10 = 0x2EED;
    const RT_DIFF10ATOM = 0x2EEE;
    const RT_DIFFTREE10 = 0x2EEC;
    const RT_DOCTOOLBARSTATES10ATOM = 0x36B1;
    const RT_DOCUMENT = 0x03E8;
    const RT_DOCUMENTATOM = 0x03E9;
    const RT_DRAWING = 0x040C;
    const RT_DRAWINGGROUP = 0x040B;
    const RT_ENDDOCUMENTATOM = 0x03EA;
    const RT_EXTERNALAVIMOVIE = 0x1006;
    const RT_EXTERNALCDAUDIO = 0x100E;
    const RT_EXTERNALCDAUDIOATOM = 0x1012;
    const RT_EXTERNALHYPERLINK = 0x0FD7;
    const RT_EXTERNALHYPERLINK9 = 0x0FE4;
    const RT_EXTERNALHYPERLINKATOM = 0x0FD3;
    const RT_EXTERNALHYPERLINKFLAGSATOM = 0x1018;
    const RT_EXTERNALMCIMOVIE = 0x1007;
    const RT_EXTERNALMEDIAATOM = 0x1004;
    const RT_EXTERNALMIDIAUDIO = 0x100D;
    const RT_EXTERNALOBJECTLIST = 0x0409;
    const RT_EXTERNALOBJECTLISTATOM = 0x040A;
    const RT_EXTERNALOBJECTREFATOM = 0x0BC1;
    const RT_EXTERNALOLECONTROL = 0x0FEE;
    const RT_EXTERNALOLECONTROLATOM = 0x0FFB;
    const RT_EXTERNALOLEEMBED = 0x0FCC;
    const RT_EXTERNALOLEEMBEDATOM = 0x0FCD;
    const RT_EXTERNALOLELINK = 0x0FCE;
    const RT_EXTERNALOLELINKATOM = 0x0FD1;
    const RT_EXTERNALOLEOBJECTATOM = 0x0FC3;
    const RT_EXTERNALOLEOBJECTSTG = 0x1011;
    const RT_EXTERNALVIDEO = 0x1005;
    const RT_EXTERNALWAVAUDIOEMBEDDED = 0x100F;
    const RT_EXTERNALWAVAUDIOEMBEDDEDATOM = 0x1013;
    const RT_EXTERNALWAVAUDIOLINK = 0x1010;
    const RT_ENVELOPEDATA9ATOM = 0x1785;
    const RT_ENVELOPEFLAGS9ATOM = 0x1784;
    const RT_ENVIRONMENT = 0x03F2;
    const RT_FONTCOLLECTION = 0x07D5;
    const RT_FONTCOLLECTION10 = 0x07D6;
    const RT_FONTEMBEDDATABLOB = 0x0FB8;
    const RT_FONTEMBEDFLAGS10ATOM = 0x32C8;
    const RT_FILTERPRIVACYFLAGS10ATOM = 0x36B0;
    const RT_FONTENTITYATOM = 0x0FB7;
    const RT_FOOTERMETACHARATOM = 0x0FFA;
    const RT_GENERICDATEMETACHARATOM = 0x0FF8;
    const RT_GRIDSPACING10ATOM = 0x040D;
    const RT_GUIDEATOM = 0x03FB;
    const RT_HANDOUT = 0x0FC9;
    const RT_HASHCODEATOM = 0x2B00;
    const RT_HEADERSFOOTERS = 0x0FD9;
    const RT_HEADERSFOOTERSATOM = 0x0FDA;
    const RT_HEADERMETACHARATOM = 0x0FF9;
    const RT_HTMLDOCINFO9ATOM = 0x177B;
    const RT_HTMLPUBLISHINFOATOM = 0x177C;
    const RT_HTMLPUBLISHINFO9 = 0x177D;
    const RT_INTERACTIVEINFO = 0x0FF2;
    const RT_INTERACTIVEINFOATOM = 0x0FF3;
    const RT_KINSOKU = 0x0FC8;
    const RT_KINSOKUATOM = 0x0FD2;
    const RT_LEVELINFOATOM = 0x2B0A;
    const RT_LINKEDSHAPE10ATOM = 0x2EE6;
    const RT_LINKEDSLIDE10ATOM = 0x2EE7;
    const RT_LIST = 0x07D0;
    const RT_MAINMASTER = 0x03F8;
    const RT_MASTERTEXTPROPATOM = 0x0FA2;
    const RT_METAFILE = 0x0FC1;
    const RT_NAMEDSHOW = 0x0411;
    const RT_NAMEDSHOWS = 0x0410;
    const RT_NAMEDSHOWSLIDESATOM = 0x0412;
    const RT_NORMALVIEWSETINFO9 = 0x0414;
    const RT_NORMALVIEWSETINFO9ATOM = 0x0415;
    const RT_NOTES= 0x03F0;
    const RT_NOTESATOM = 0x03F1;
    const RT_NOTESTEXTVIEWINFO9 = 0x0413;
    const RT_OUTLINETEXTPROPS9 = 0x0FAE;
    const RT_OUTLINETEXTPROPS10 = 0x0FB3;
    const RT_OUTLINETEXTPROPS11 = 0x0FB5;
    const RT_OUTLINETEXTPROPSHEADER9ATOM = 0x0FAF;
    const RT_OUTLINETEXTREFATOM = 0x0F9E;
    const RT_OUTLINEVIEWINFO = 0x0407;
    const RT_PERSISTDIRECTORYATOM = 0x1772;
    const RT_PARABUILD = 0x2B08;
    const RT_PARABUILDATOM = 0x2B09;
    const RT_PHOTOALBUMINFO10ATOM = 0x36B2;
    const RT_PLACEHOLDERATOM = 0x0BC3;
    const RT_PRESENTATIONADVISORFLAGS9ATOM = 0x177A;
    const RT_PRINTOPTIONSATOM = 0x1770;
    const RT_PROGBINARYTAG = 0x138A;
    const RT_PROGSTRINGTAG = 0x1389;
    const RT_PROGTAGS = 0x1388;
    const RT_RECOLORINFOATOM = 0x0FE7;
    const RT_RTFDATETIMEMETACHARATOM = 0x1015;
    const RT_ROUNDTRIPANIMATIONATOM12ATOM = 0x2B0B;
    const RT_ROUNDTRIPANIMATIONHASHATOM12ATOM = 0x2B0D;
    const RT_ROUNDTRIPCOLORMAPPING12ATOM = 0x040F;
    const RT_ROUNDTRIPCOMPOSITEMASTERID12ATOM = 0x041D;
    const RT_ROUNDTRIPCONTENTMASTERID12ATOM = 0x0422;
    const RT_ROUNDTRIPCONTENTMASTERINFO12ATOM = 0x041E;
    const RT_ROUNDTRIPCUSTOMTABLESTYLES12ATOM = 0x0428;
    const RT_ROUNDTRIPDOCFLAGS12ATOM = 0x0425;
    const RT_ROUNDTRIPHEADERFOOTERDEFAULTS12ATOM = 0x0424;
    const RT_ROUNDTRIPHFPLACEHOLDER12ATOM = 0x0420;
    const RT_ROUNDTRIPNEWPLACEHOLDERID12ATOM = 0x0BDD;
    const RT_ROUNDTRIPNOTESMASTERTEXTSTYLES12ATOM = 0x0427;
    const RT_ROUNDTRIPOARTTEXTSTYLES12ATOM = 0x0423;
    const RT_ROUNDTRIPORIGINALMAINMASTERID12ATOM = 0x041C;
    const RT_ROUNDTRIPSHAPECHECKSUMFORCL12ATOM = 0x0426;
    const RT_ROUNDTRIPSHAPEID12ATOM = 0x041F;
    const RT_ROUNDTRIPSLIDESYNCINFO12 = 0x3714;
    const RT_ROUNDTRIPSLIDESYNCINFOATOM12 = 0x3715;
    const RT_ROUNDTRIPTHEME12ATOM = 0x040E;
    const RT_SHAPEATOM = 0x0BDB;
    const RT_SHAPEFLAGS10ATOM = 0x0BDC;
    const RT_SLIDE = 0x03EE;
    const RT_SLIDEATOM = 0x03EF;
    const RT_SLIDEFLAGS10ATOM = 0x2EEA;
    const RT_SLIDELISTENTRY10ATOM = 0x2EF0;
    const RT_SLIDELISTTABLE10 = 0x2EF1;
    const RT_SLIDELISTWITHTEXT = 0x0FF0;
    const RT_SLIDELISTTABLESIZE10ATOM = 0x2EEF;
    const RT_SLIDENUMBERMETACHARATOM = 0x0FD8;
    const RT_SLIDEPERSISTATOM = 0x03F3;
    const RT_SLIDESHOWDOCINFOATOM = 0x0401;
    const RT_SLIDESHOWSLIDEINFOATOM = 0x03F9;
    const RT_SLIDETIME10ATOM = 0x2EEB;
    const RT_SLIDEVIEWINFO = 0x03FA;
    const RT_SLIDEVIEWINFOATOM = 0x03FE;
    const RT_SMARTTAGSTORE11CONTAINER = 0x36B3;
    const RT_SOUND = 0x07E6;
    const RT_SOUNDCOLLECTION = 0x07E4;
    const RT_SOUNDCOLLECTIONATOM = 0x07E5;
    const RT_SOUNDDATABLOB = 0x07E7;
    const RT_SORTERVIEWINFO = 0x0408;
    const RT_STYLETEXTPROPATOM = 0x0FA1;
    const RT_STYLETEXTPROP10ATOM = 0x0FB1;
    const RT_STYLETEXTPROP11ATOM = 0x0FB6;
    const RT_STYLETEXTPROP9ATOM = 0x0FAC;
    const RT_SUMMARY = 0x0402;
    const RT_TEXTBOOKMARKATOM = 0x0FA7;
    const RT_TEXTBYTESATOM = 0x0FA8;
    const RT_TEXTCHARFORMATEXCEPTIONATOM = 0x0FA4;
    const RT_TEXTCHARSATOM = 0x0FA0;
    const RT_TEXTDEFAULTS10ATOM = 0x0FB4;
    const RT_TEXTDEFAULTS9ATOM = 0x0FB0;
    const RT_TEXTHEADERATOM = 0x0F9F;
    const RT_TEXTINTERACTIVEINFOATOM = 0x0FDF;
    const RT_TEXTMASTERSTYLEATOM = 0x0FA3;
    const RT_TEXTMASTERSTYLE10ATOM = 0x0FB2;
    const RT_TEXTMASTERSTYLE9ATOM = 0x0FAD;
    const RT_TEXTPARAGRAPHFORMATEXCEPTIONATOM = 0x0FA5;
    const RT_TEXTRULERATOM = 0x0FA6;
    const RT_TEXTSPECIALINFOATOM = 0x0FAA;
    const RT_TEXTSPECIALINFODEFAULTATOM = 0x0FA9;
    const RT_TIMEANIMATEBEHAVIOR = 0xF134;
    const RT_TIMEANIMATEBEHAVIORCONTAINER = 0xF12B;
    const RT_TIMEANIMATIONVALUE = 0xF143;
    const RT_TIMEANIMATIONVALUELIST = 0xF13F;
    const RT_TIMEBEHAVIOR = 0xF133;
    const RT_TIMEBEHAVIORCONTAINER = 0xF12A;
    const RT_TIMECOLORBEHAVIOR = 0xF135;
    const RT_TIMECOLORBEHAVIORCONTAINER = 0xF12C;
    const RT_TIMECLIENTVISUALELEMENT = 0xF13C;
    const RT_TIMECOMMANDBEHAVIOR = 0xF13B;
    const RT_TIMECOMMANDBEHAVIORCONTAINER = 0xF132;
    const RT_TIMECONDITION = 0xF128;
    const RT_TIMECONDITIONCONTAINER = 0xF125;
    const RT_TIMEEFFECTBEHAVIOR = 0xF136;
    const RT_TIMEEFFECTBEHAVIORCONTAINER = 0xF12D;
    const RT_TIMEEXTTIMENODECONTAINER = 0xF144;
    const RT_TIMEITERATEDATA = 0xF140;
    const RT_TIMEMODIFIER = 0xF129;
    const RT_TIMEMOTIONBEHAVIOR = 0xF137;
    const RT_TIMEMOTIONBEHAVIORCONTAINER = 0xF12E;
    const RT_TIMENODE = 0xF127;
    const RT_TIMEPROPERTYLIST = 0xF13D;
    const RT_TIMEROTATIONBEHAVIOR = 0xF138;
    const RT_TIMEROTATIONBEHAVIORCONTAINER = 0xF12F;
    const RT_TIMESCALEBEHAVIOR = 0xF139;
    const RT_TIMESCALEBEHAVIORCONTAINER = 0xF130;
    const RT_TIMESEQUENCEDATA = 0xF141;
    const RT_TIMESETBEHAVIOR = 0xF13A;
    const RT_TIMESETBEHAVIORCONTAINER = 0xF131;
    const RT_TIMESUBEFFECTCONTAINER = 0xF145;
    const RT_TIMEVARIANT = 0xF142;
    const RT_TIMEVARIANTLIST = 0xF13E;
    const RT_USEREDITATOM = 0x0FF5;
    const RT_VBAINFO = 0x03FF;
    const RT_VBAINFOATOM = 0x0400;
    const RT_VIEWINFOATOM = 0x03FD;
    const RT_VISUALPAGEATOM = 0x2B01;
    const RT_VISUALSHAPEATOM = 0x2AFB;

    /* @var \Cryptodira\PhpOffice\OLEDocument */
    private $oleDocument;
    private $streamId;
    private $stream;
    private $summaryStreamId;
    private $docinfoStreamId;
    private $summaryInfo;
    private $docSummaryInfo;
    private $offsetToCurrentEdit;
    private $persistDirectory;
    private $text = '';
    private $slideIds = [];
    private $masterIds = [];
    private $notesIds = [];
    private $slides = [];
    private $masters = [];
    private $notes = [];

    public function __construct($file)
    {
        $this->oleDocument = new OLEDocument();
        if (is_string($file)) {
            $this->oleDocument->CreateFromFile($file);
        }
        elseif (is_resource($file) && get_resource_type($file) == 'stream') {
            $this->oleDocument->CreateFromStream($file);
        }
        else {
            throw new \Exception("Must pass a file name or stream to PPT Document constructor");
        }

        $this->streamId = $this->oleDocument->FindStreamByName('PowerPoint Document');
        $this->summaryStreamId = $this->oleDocument->FindStreamByName(chr(5) . 'SummaryInformation');
        $this->docinfoStreamId = $this->oleDocument->FindStreamByName(chr(5) . 'DocumentSummaryInformation');
        $this->docSummaryInfo = null;
        $this->summaryInfo = null;
        $this->stream = new OLEStreamReader($this->oleDocument, $this->streamId);

        $this->readCurrentUserStream();
        $this->readDocumentStream();
    }

    public function getDocumentSummary()
    {
        if (is_null($this->docSummaryInfo)) {
            $this->docSummaryInfo = new PropertyBag($this->oleDocument->Read($this->docinfoStreamId));
        }
        return $this->docSummaryInfo;
    }

    public function getSummaryInfo()
    {
        if (is_null($this->summaryinfo)) {
            $this->summaryInfo = new PropertyBag($this->oleDocument->Read($this->summaryStreamId));
        }
        return $this->summaryInfo;
    }

    private function readRecordHeader($stream=null)
    {
        if (!$stream) {
            $stream = $this->stream;
        }

        $rec = $stream->readUint2();
        $recType = $stream->readUint2();
        $recLen = $stream->readUint4();
        return array(
            'recVer' => ($rec >> 0) & bindec('1111'),
            'recInstance' => ($rec >> 4) & bindec('111111111111'),
            'recType' => $recType,
            'recLen' => $recLen,
        );
    }

    private function skipRecord($stream=null)
    {
        if (!$stream) {
            $stream = $this->stream;
        }

        $h = $this->readRecordHeader($stream);
        $stream->seek($h[recLen], SEEK_CUR);
    }

    private function readCurrentUserStream()
    {
        $currentUserStreamId = $this->oleDocument->FindStreamByName('Current User');
        $stream = new OLEStreamReader($this->oleDocument, $currentUserStreamId);
        $h = $this->readRecordHeader($stream);
        if ($h['recType'] != self::RT_CURRENTUSERATOM) {
            throw new \Exception('Unexpected record type in User Stream');
        }

        $size = $stream->readUint4();
        $headerToken = $stream->readUint4();
        $this->offsetToCurrentEdit = $stream->readUint4();
    }

    private function readUserEditAtom($offset)
    {
        $this->stream->seek($offset);
        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_USEREDITATOM) {
            throw new \Exception('Wrong record type for UserEditAtom');
        }
        $res = array();
        $res['lastSlideRef'] = $this->stream->readUint4();
        $res['version'] = $this->stream->readUint4();
        $res['offsetLastEdit'] = $this->stream->readUint4();
        $res['offsetPersistDirectory'] = $this->stream->readUint4();
        $res['docPersistIdRef'] = $this->stream->readUint4();
        return $res;
    }

    private function readPersistDirectoryAtom($offset)
    {
        $this->stream->seek($offset);
        $h = $this->readRecordHeader();
        $b =0;
        $res = array();
        while ($b < $h['recLen']) {
            $v = $this->stream->readUint4();
            $persistID = ($v & 0xFFFFF);
            $cPersist = ($v >> 20) & 0xFFF;
            $offsets = unpack('V' . $cPersist, $this->stream->read(4 * $cPersist));

            // can't use array_merge for this because it renumbers numeric keys
            for ($i=0; $i < $cPersist; $i++) {
                $res[$persistID+$i] = $offsets[$i+1];
            }
            $b += 4 * ($cPersist+1);
        }
        return $res;
    }

    /* Slides, Masters */

    private function readSlideAtom()
    {
        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_SLIDEATOM || $h['recVer'] != 0x2) {
            throw new \Exception('Wrong recType for DocumentContainer');
        }
        $slide = unpack('V1geom/c8rgPlaceholderTypes/V1masterIdRef/' .
                'V1notesIdRef/v1slideFlags/v1unused', $this->stream->read(24));
        return $slide;
    }

    private function readHeaderFooter($instance)
    {
        $res = array(
            'customDate' => '',
            'headerText' => '',
            'footerText' => '',
        );

        // hfAtom
        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_HEADERSFOOTERSATOM) {
            throw new \Exception('Wrong recType for HeaderFooterAtom');
        }
        $formatID = $this->stream->readUint2();
        $hfFlags = $this->stream->readUint2();
        $fHasDate = $hfFlags & 0x1;
        $fHasTodayDate = $hfFlags & 0x2;
        $fHasUserDate = $hfFlags & 0x4;
        $fHasSlideNumber = $hfFlags & 0x8;
        $fHasHeader = $hfFlags & 0x10;
        $fHasFooter = $hfFlags & 0x20;

         // userDate
        if ($fHasUserDate) {
            $h = $this->readRecordHeader();
            if ($h['recType'] != self::RT_CSTRING) {
                throw new \Exception('Wrong recType for user date');
            }
            $res['customDate'] = mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'UTF-16LE');
        }

        // headerAtom
        if ($instance == 0x04 && $fHasHeader) {
            $h = $this->readRecordHeader();
            if ($h['recType'] != self::RT_CSTRING) {
                throw new \Exception('Wrong recType for header string');
            }
            $res['headerText'] = mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'UTF-16LE');
        }

        // footerAtom
        if ($fHasFooter) {
            $h = $this->readRecordHeader();
            if ($h['recType'] != self::RT_CSTRING) {
                throw new \Exception('Wrong recType for footer string');
            }
            $res['footerText'] = mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'UTF-16LE');
        }

        return $res;
    }

    private function readPP10BinaryTagExtension(&$shape)
    {
        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_BINARYTAGDATABLOB) {
            throw new \Exception('Wrong record type for binary tag blob');
        }

        $b = $h['recLen'];
        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case self::RT_TEXTMASTERSTYLE10ATOM:
                case self::RT_COMMENT10:
                case self::RT_LINKEDSLIDE10ATOM:
                case self::RT_LINKEDSHAPE10ATOM:
                case self::RT_SLIDEFLAGS10ATOM:
                case self::RT_SLIDETIME10ATOM:
                case self::RT_HASHCODEATOM:
                case self::RT_TIMEEXTTIMENODECONTAINER:
                case self::RT_BUILDLIST:
                default:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
            }
            $b -= (8 + $h['recLen']);
        }

    }

    private function readProgTagsContainer($recLen, &$shape)
    {
        $b = $recLen;
        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case self::RT_PROGSTRINGTAG:
                case self::RT_PROGBINARYTAG:
                    $th = $this->readRecordHeader();
                    if ($th['recType'] != self::RT_CSTRING) {
                        throw new \Exception('ProgTag name must be CSTRING');
                    }
                    $name = mb_convert_encoding($this->stream->read($th['recLen']), 'UTF-8', 'UTF-16LE') . "\r";
                    if ($h['recType'] == self::RT_PROGSTRINGTAG) {
                        $th = $this->readRecordHeader();
                        if ($th['recType'] != self::RT_CSTRING) {
                         throw new \Exception('ProgTag value must be CSTRING');
                            }
                        $value = mb_convert_encoding($this->stream->read($th['recLen']), 'UTF-8', 'UTF-16LE') . "\r";
                    }
                    elseif ($name == '___PPT10') {
                        $this->readPP10BinaryTagExtension($shape);
                    }
                    else {
                        $this->stream->seek($h['recLen'] - (8+$th['recLen']), SEEK_CUR);
                    }
                    break;
                default:
                    throw new \Exception('This should never happen.');
            }
            $b -= (8 + $h['recLen']);
        }
    }

    private function readOfficeArtClientData($recLen, &$shape)
    {
        $b = $recLen;
        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case self::RT_SHAPEATOM:
                case self::RT_SHAPEFLAGS10ATOM:
                    $flags = $this->stream->readUint1();
                    break;
                case self::RT_EXTERNALOBJECTREFATOM:
                    $exObjID = $this->stream->readUint4();
                    break;
                case self::RT_ANIMATIONINFO:
                case self::RT_INTERACTIVEINFO:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
                case self::RT_PLACEHOLDERATOM:
                    $position = $this->stream->readUint4();
                    $placementId = $this->stream->readUint1();
                    $size = $this->stream->readUint1();
                    $this->stream->seek(2, SEEK_CUR);
                    break;
                case self::RT_PROGTAGS:
                    $this->readProgTagsContainer($h['recLen'], $shape);
                    break;
                case self::RT_RECOLORINFOATOM:
                default:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
            }
            $b -= (8 + $h['recLen']);
        }
    }

    private function readTextClientDataSubContainerOrAtom($recLen)
    {
        $text = '';

        $b = $recLen;
        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case self::RT_OUTLINETEXTREFATOM:
                case self::RT_TEXTHEADERATOM:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
                case self::RT_TEXTCHARSATOM:
                    $text .= mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'UTF-16LE') . "\r";
                    break;
                case self::RT_TEXTBYTESATOM:
                    $text .= mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'ASCII') . "\r";
                    break;
                case self::RT_STYLETEXTPROPATOM:
                case self::RT_SLIDENUMBERMETACHARATOM:
                case self::RT_DATETIMEMETACHARATOM:
                case self::RT_GENERICDATEMETACHARATOM:
                case self::RT_HEADERMETACHARATOM:
                case self::RT_FOOTERMETACHARATOM:
                case self::RT_RTFDATETIMEMETACHARATOM:
                case self::RT_TEXTBOOKMARKATOM:
                case self::RT_TEXTSPECIALINFOATOM:
                case self::RT_INTERACTIVEINFO:
                case self::RT_TEXTINTERACTIVEINFOATOM:
                case self::RT_TEXTRULERATOM:
                case self::RT_MASTERTEXTPROPATOM:
                default:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
            }
            $b -= (8 + $h['recLen']);
        }
        return $text;
    }

    private function readOptions($numProps, &$shape)
    {
        $dataOffset = $this->stream->tell() + $numProps * 6;
        for ($i = 0; $i < $numProps; $i++) {
            $opid = $this->stream->readUint2();
            $fBid = $opid & 0x4000;
            $fComplex = $opid & 0x8000;
            $opid &= 0x3FFF;
            $op = $this->stream->readUint4();
            if ($op && $fComplex) {
                $pos = $this->stream->tell();
                $this->stream->seek($dataOffset);
                $s = $this->stream->read($op);
                $dataOffset += $op;
                $this->stream->seek($pos);
            }
        }
        // by the end, this should point to the end of the property array + data
        $this->stream->seek($dataOffset);
    }

    private function readOfficeArtSpContainer($recLen)
    {
        $shape = array();

        $b = $recLen;
        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case 0xF009: // shapeGroup
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
                case 0xF00A: // shapeProp
                    $shape['id'] = $this->stream->readUint4();
                    $shapeProps = $this->stream->readUint4();
                    $fGroup = $shapeProps & 0x1;
                    $fChild = $shapeProps & 0x2;
                    $fPatriarch = $shapeProps & 0x4;
                    $fDeleted = $shapeProps & 0x8;
                    $fOleShape = $shapeProps & 0x10;
                    $fHaveMaster = $shapeProps & 0x20;
                    $fFlipH = $shapeProps & 0x40;
                    $fFlipV = $shapeProps & 0x80;
                    $fConnector = $shapeProps & 0x100;
                    $fHaveAnchor = $shapeProps & 0x200;
                    $fBackground = $shapeProps & 0x400;
                    $fHaveSpt = $shapeProps & 0x800;
                    break;
                case 0xF00B: // shapePrimaryOptions
                case 0xF121: // shapeSecondaryOptions1
                case 0xF122: // shapeTertiaryOptions1
                    $pos = $this->stream->tell();
                    $this->readOptions($h['recInstance'], $shape);
                    $this->stream->seek($pos + $h['recLen']);
                    break;
                case 0xF11D: // deletedShape
                case 0xF00F: // childAnchor
                case 0xF010: // clientAnchor
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
                case 0xF011: // clientData
                    $this->readOfficeArtClientData($h['recLen'], $shape);
                    break;
                case 0xF00D: // clientText
                    $shape['clientText'] = $this->readTextClientDataSubContainerOrAtom($h['recLen']);
                    break;
                default:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
            }
            $b -= (8 + $h['recLen']);
        }

        return $shape;
    }

    private function readOfficeArtSpgrContainer($recLen)
    {
        $group = array();
        $b = $recLen;
        while ($b > 0) {
            $h = $this->readRecordHeader();
            if ($h['recType'] == 0xF003) {
                $group[] = $this->readOfficeArtSpgrContainer($h['recLen']);
            }
            elseif ($h['recType'] == 0xF004) {
                $group[] = $this->readOfficeArtSpContainer($h['recLen']);
            }
            else {
                break;
            }
            $b -= (8 + $h['recLen']);
        }

        return $group;
    }

    private function readOfficeArtDgContainer(&$slide) // OfficeArtDgContainer
    {
        $h = $this->readRecordHeader();
        if ($h['recType'] != 0xF002) {
            throw new \Exception('Wrong recType for drawing container');
        }
        $b = $h['recLen'];

        // OfficeArtFDG
        $h = $this->readRecordHeader();
        if ($h['recType'] != 0xF008) {
            throw new \Exception('Wrong recType for drawing container info');
        }
        $csp = $this->stream->readUint4();     // number of shapes in the drawing
        $spidCur = $this->stream->readUint4(); // id of last shapein drawing
        $b -= 16;
        $text = '';

        $shapeGroup = null;
        $shape = null;
        $deleted = array();
        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case 0xF003: // groupShape
                    if (!$shapeGroup && !$shape) {
                        $shapeGroup = $this->readOfficeArtSpgrContainer($h['recLen']);
                        foreach ($shapeGroup as $s) {
                            if (isset($s['clientText'])) {
                                $text .= $s['clientText'] . "\r";
                            }
                        }
                    }
                    else {
                        $deleted[] = $this->readOfficeArtSpgrContainer($h['recLen']);
                    }
                    break;
                case 0xF004: // shape
                    if (!$shape) {
                        $shape = $this->readOfficeArtSpContainer($h['recLen']);
                        if (isset($shape['clientText'])) {
                            $text .= $shape['clientText'] . "\r";
                        }
                    }
                    else {
                        $deleted[] = $this->readOfficeArtSpContainer($h['recLen']);
                    }
                    break;
                case 0xF118: // OfficeArtFRITContainer regroupItems
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
                case 0xF005: // OfficeArtSolverContainer solver1, solver2
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
                default:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
            }
            $b -= (8+$h['recLen']);
        }

        $slide['text'] = $text;
        $slide['drawing'] = ['shapeGroup' => $shapeGroup, 'shape' => $shape, 'deleted' => $deleted];
    }

    private function readSlideContainer($persistId)
    {
        $this->stream->seek($this->persistDirectory[$persistId]);
        $h = $this->readRecordHeader();
        if (($h['recType'] != self::RT_SLIDE && $h['recType'] != self::RT_MAINMASTER) || $h['recVer'] != 0xF) {
            throw new \Exception('Wrong recType for Slide or MainMaster Container');
        }
        $slide = array_merge($this->readSlideAtom(),
            ['name' => '', 'text' => '', 'footerText' => '']);

        $b = $h['recLen'] - 32;
        while ($b > 0) {
            $h1 = $this->readRecordHeader();
            switch ($h1['recType']) {
                case self::RT_CSTRING: // slideNameAtom
                    if ($h1['recInstance'] == 3) {
                        $slide['name'] = mb_convert_encoding($this->stream->read($h1['recLen']), 'UTF-8', 'UTF-16LE');
                    }
                    elseif ($h1['recInstance'] == 2) {
                        $slide['templateName'] = mb_convert_encoding($this->stream->read($h1['recLen']), 'UTF-8', 'UTF-16LE');
                    }
                    else {
                        $this->stream->seek($h1['recLen'], SEEK_CUR);
                    }
                    break;
                case self::RT_DRAWING:
                    $this->readOfficeArtDgContainer($slide);
                    break;
                case self::RT_HEADERSFOOTERS: //perSlideHFContainer
                    $hf = $this->readHeaderFooter($h1['recInstance']);
                    $slide['customDate'] = $hf['customDate'];
                    $slide['footerText'] = $hf['footerText'];
                    break;
                case self::RT_PROGTAGS:
                    $this->readProgTagsContainer($h1['recLen'], $slide);
                    break;
                case self::RT_SLIDESHOWSLIDEINFOATOM:
                case self::RT_TEXTMASTERSTYLEATOM:
                case self::RT_ROUNDTRIPSLIDESYNCINFO12:
                case self::RT_COLORSCHEMEATOM:
                default:
                    $this->stream->seek($h1['recLen'], SEEK_CUR);
            }
            $b -= (8 + $h1['recLen']);
        }

        return $slide;
    }

    private function readSlides($idList, &$itemList)
    {
        foreach ($idList as $id => $persistId) {
            $itemList[$id] = $this->readSlideContainer($persistId);
        }
    }

    /* Notes */

    private function readNotesContainer($persistID)
    {
        if ($this->stream->seek($this->persistDirectory[$persistID]) != 0) {
            return null;
        }

        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_NOTES || $h['recVer'] != 0xF) {
            throw new \Exception('Invalid record type for Notes Container');
        }
        $b = $h['recLen'];

        // Read the NotesAtom record
        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_NOTESATOM || $h['recVer'] != 0x1) {
            throw new \Exception('Mising Notes Atom in Notes Container');
        }
        $note = array('name' => '', 'text' => '', 'footerText' => '');
        $slideId = $this->stream->readUint4();
        if ($slideId && isset($this->slides[$slideId])) {
            $note['slide'] = $this->slides[$slideId];
        }
        $note['slideFlags'] = $this->stream->readUint4();
        $this->stream->seek(2, SEEK_CUR);
        $b -= 16;

        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case self::RT_CSTRING: // slideNameAtom
                    $note['name'] = mb_convert_encoding($this->stream->read($h1['recLen']), 'UTF-8', 'UTF-16LE');
                    break;
                case self::RT_DRAWING:
                    $this->readOfficeArtDgContainer($note);
                    break;
                case self::RT_PROGTAGS:
                    $this->readProgTagsContainer($h['recLen'], $slide);
                    break;
                case self::RT_COLORSCHEMEATOM:
                case self::RT_ROUNDTRIPTHEME12ATOM:
                case self::RT_ROUNDTRIPCOLORMAPPING12ATOM:
                case self::RT_ROUNDTRIPNOTESMASTERTEXTSTYLES12ATOM:
                default:
                    $this->seek($h['recLen'], SEEK_CUR);
                    break;
            }
            $b -= (8 + $h['recLen']);
        }

        return $note;
    }

    private function ReadNotes()
    {
        foreach ($this->notesIds as $id => $persistId) {
            $this->notes[$id] = $this->readNotesContainer($persistId);
        }
    }

    /* Handout */

    private function readHandoutContainer($persistId)
    {
        if ($this->stream->seek($this->persistDirectory[$persistId]) != 0) {
            return null;
        }

        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_HANDOUT || $h['recVer'] != 0xF) {
            throw new \Exception('Invalid record type for Handout Container');
        }
        $b = $h['recLen'];
        $handout = array('name' => '', 'text' => '', 'footerText' => '');

        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case self::RT_CSTRING:
                    $handout['name'] = mb_convert_encoding($this->stream->read($h1['recLen']), 'UTF-8', 'UTF-16LE');
                    break;
                case self::RT_DRAWING:
                    $this->readOfficeArtDgContainer($handout);
                    break;
                case self::RT_PROGTAGS:
                    $this->readProgTagsContainer($h['recLen'], $slide);
                    break;
                case self::RT_COLORSCHEMEATOM:
                case self::RT_ROUNDTRIPTHEME12ATOM:
                case self::RT_ROUNDTRIPCOLORMAPPING12ATOM:
                case self::RT_ROUNDTRIPNOTESMASTERTEXTSTYLES12ATOM:
                default:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
            }
            $b -= (8 + $h['recLen']);
        }

        return $handout;
    }

    /* Document Container */

    private function readDocumentAtom()
    {
        $atom = unpack('V1SlideSizeX/V1SlideSizeY/V1NotesSizeX/V1NotesSizeY/' .
            'P1ServerZoom/V1notesMasterPersistIdRef/V1handoutMasterPersistIdRef/' .
            'v1firstSlideNumber/v1slideSizeType/' .
            'c1fSaveWithFonts/c1fOmitTitlePlace/c1fRightToLeft/c1fShowComments',
            $this->stream->read(40));
        $this->notesPersistID = $atom['notesMasterPersistIdRef'];
        $this->handoutPersistID = $atom['handoutMasterPersistIdRef'];
        return $atom;
    }

    private function readSlideList($recLen, $instance)
    {
        $b = $recLen;
        while ($b > 0) {
            $h = $this->readRecordHeader();
            switch ($h['recType']) {
                case self::RT_SLIDEPERSISTATOM:
                    $slidePersistId = $this->stream->readUint4();
                    $this->stream->seek(8, SEEK_CUR); // A, B, C, reserved2, cTexts
                    $slideId = $this->stream->readUint4();
                    $this->stream->seek(4, SEEK_CUR); // reserved3

                    switch ($instance) {
                        case 0:
                            $this->slideIds[$slideId] = $slidePersistId;
                            break;
                        case 1:
                            $this->masterIds[$slideId] = $slidePersistId;
                            break;
                        case 2:
                            $this->notesIds[$slideId] = $slidePersistId;
                            break;
                    }

                    break;
                case self::RT_TEXTCHARSATOM:
                    $s = mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'UTF-16LE') . '\n';
                    break;
                case self::RT_TEXTBYTESATOM:
                    $s = mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'ASCII') . '\n';
                    break;
                default:
                    $this->stream->seek($h['recLen'], SEEK_CUR);
                    break;
            }
            $b -= (8 + $h['recLen']);
        }
    }

    private function readDocumentContainer($offset)
    {
        $this->stream->seek($offset);
        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_DOCUMENT || $h['recVer'] != 0xF) {
            throw new \Exception('Wrong recType for DocumentContainer');
        }

        $b = $h['recLen'];
        while ($b > 0) {
            $h1 = $this->readRecordHeader();
            switch ($h1['recType']) {
                case self::RT_DOCUMENTATOM:
                    $this->readDocumentAtom();
                    break;
                case self::RT_EXTERNALOBJECTLIST:
                    $this->stream->seek($h1['recLen'], SEEK_CUR);
                    break;
                case self::RT_ENVIRONMENT:
                case self::RT_SOUNDCOLLECTION:
                case self::RT_DRAWINGGROUP:
                case self::RT_LIST: // docInfoList
                    $this->stream->seek($h1['recLen'], SEEK_CUR);
                    break;
                case self::RT_HEADERSFOOTERS: // slideHF (instance=3), notesHF (instance=4)
                    $hf = $this->readHeaderFooter($h1['recInstance']);
                    if ($h1['recInstance'] == 0x03) {
                        $this->customDate = $hf['customDate'];
                        $this->slideFooter = $hf['footerText'];
                    }
                    else {
                        $this->notesDate = $hf['customDate'];
                        $this->notesHeader = $hf['headerText'];
                        $this->notesFooter = $hf['footerText'];
                    }
                    break;
                case self::RT_SLIDELISTWITHTEXT: // slides (instance=0), masters (instance=1), notes (instance=2)
                    $this->readSlideList($h1['recLen'], $h1['recInstance']);
                    break;
                case self::RT_SLIDESHOWDOCINFOATOM:
                case self::RT_NAMEDSHOWS:
                case self::RT_SUMMARY:
                case self::RT_DOCROUTINGSLIPATOM:
                case self::RT_PRINTOPTIONSATOM:
                case self::RT_ROUNDTRIPCUSTOMTABLESTYLES12ATOM:
                case self::RT_ENDDOCUMENTATOM:
                default:
                    $this->stream->seek($h1['recLen'], SEEK_CUR);
                    break;
            }
            $b -= (8+$h1['recLen']);
        }
    }

    private function readDocumentStream()
    {
        $userEdit = $this->readUserEditAtom($this->offsetToCurrentEdit);
        $docPersistID = $userEdit['docPersistIdRef'];
        $atoms = array();
        do {
            array_unshift($atoms, $this->readPersistDirectoryAtom($userEdit['offsetPersistDirectory']));

            if ($userEdit['offsetLastEdit'] != 0) {
                $userEdit = $this->readUserEditAtom($userEdit['offsetLastEdit']);
            }
        } while ($userEdit['offsetLastEdit'] != 0);


        foreach ($atoms as $a) {
            if (!$this->persistDirectory) {
                $this->persistDirectory = $a;
            }
            else {
                // can't use array_merge for this because it renumbers numeric keys
                foreach ($a as $k => $v) {
                    $this->persistDirectory[$k] = $v;
                }
            }
        }

        $this->readDocumentContainer($this->persistDirectory[$docPersistID]);

        if ($this->notesPersistID) {
            $this->notesMaster = $this->readNotesContainer($this->notesPersistID);
        }

        if ($this->handoutPersistID) {
            $this->handoutMaster = $this->readHandoutContainer($this->handoutPersistID);
        }

        $this->readSlides($this->masterIds, $this->masters);
        $this->readSlides($this->slideIds, $this->slides);
        $this->readNotes();

        //$this->readExternalObjectList();
    }

    public function getText()
    {
        if (!$this->text) {
            $this->text = '';
            foreach ($this->masters as $slide) {
                $this->text .= ($slide['name'] ? $slide['name'] . "\n\n" : '')
                    . ($slide['templateName'] ? $slide['templateName'] . "\n\n" : '')
                    . $slide['text'] . "\n\n"
                    . ($slide['footerText'] ? $slide['footerText'] . "\n\n" : '');
            }

            foreach ($this->slides as $slide) {
                $this->text .= ($slide['name'] ? $slide['name'] . "\n\n" : '')
                    . $slide['text'] . "\n\n"
                    . ($slide['footerText'] ? $slide['footerText'] . "\n\n" : '');
            }

            if ($this->slideFooter) {
                $this->text .= $this->slideFooter . "\n\n";
            }

            foreach ($this->notes as $note) {
                $this->text .= ($note['name'] ? $note['name'] . "\n\n" : '')
                    . ($note['headerText'] ? $note['headerText'] . "\n\n" : '')
                    . $note['text'] . "\n\n"
                    . ($note['footerText'] ? $note['footerText'] . "\n\n" : '');
            }
        }
        return $this->text;
    }
}
