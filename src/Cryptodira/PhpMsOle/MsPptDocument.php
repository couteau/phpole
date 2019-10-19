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
    private $text = '';
    
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

        $this->streamId = $this->oleDocument->GetDocumentStream();
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
        $stream = new OLEStreamReader($currentUserStreamId);
        $h = $this->readRecordHeader($stream);
        if ($h['recType'] != self::RT_CURRENTUSERATOM) {
            throw new Exception('Unexpected record type in User Stream');
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
            throw new Exception('Wrong record type for UserEditAtom');
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
            $persistID = ($v >> 12);
            $cPersist = $v & bindec('111111111111');
            $offsets = unpack('V' . cPersist, $this->stream->read(4 * cPersist));
            
            // can't use array_merge for this because it renumbers numeric keys
            for ($i=0; $i < $cPersist; $i++) {
                $res[$persistID+$i] = $offsets[$i+1];
            }            
            $this->appendPersistDirectory($res, $persistID, $cPersist, $offsets);
            $b += 4 * (cPersist+1);
        }
        return $res;
    }
    
    private function readDocumentAtom()
    {
        $atom = unpack('V1SlideSizeX/V1SlideSizeY/V1NotesSizeX/V1NotesSizeY/' . 
                'P1ServerZoom/V1notesMasterPersistIdRef/V1handoutMasterPersistIdRef/' . 
                'v1firstSlideNumber/v1slideSizeType/' . 
                'c1fSaveWithFonts/c1fOmitTitlePlace/c1fRightToLeft/c1fShowComments');
        $this->notesPersistID = $atom['notesMasterPersistIDRef'];
        $this->handoutPersistID = $atom['handoutMasterPersistIDRef'];
        return $atom;
    }
    
    private function readExObjectList()
    {
        $this->skiprecord();
    }
    
    private function readDocumentTextInfo()
    {
        $this->skiprecord();
    }
    
    private function readSoundCollection()
    {
        $this->skiprecord();
    }
    
    private function readDrawingGroup()
    {
        $this->skiprecord();
    }
    
    private function readMasterList()
    {
        $this->skiprecord();
    }
    
    private function readDocInfoList()
    {
        $this->skiprecord();
    }
    
    private function readSlideHF()
    {
        $this->skiprecord();
    }
    
    private function readNotesHF()
    {
        $this->skiprecord();
    }
    
    private function readSlideListWithTextSubContainerOrAtom()
    {
        $h = $this->readRecordHeader();
        switch ($h['recType']) {
            case self::RT_TEXTCHARSATOM:
                $s = mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'UTF-16LE') . '\n';
                $this->text .= $s;
                break;
            case self::RT_TEXTBYTESATOM:
                $s = mb_convert_encoding($this->stream->read($h['recLen']), 'UTF-8', 'ASCII') . '\n';
                $this->text .= $s;
                break;
            default:
                $this->stream->seek($h['recLen'], SEEK_CUR);
                break;
        }
        return 8+$h['recLen'];
    }
    
    private function readSlideList($recLen) 
    {
        $b = $recLen;
        while ($b > 0) {
            $b -= $this->readSlideListWithTextSubContainerOrAtom();
        }
    }
    
    
    private function readDocumentContainer($offset)
    {
        $this->stream->seek($offset);
        $h = $this->readRecordHeader();
        if ($h['recType'] != self::RT_DOCUMENT || $h['recVer'] != 0xF) {
            throw new Exception('Wrong recType for DocumentContainer');
        }
        
        $b = $h['recLen'];
        while ($b > 0) {
            $h1 = $this->readRecordHeader();
            switch ($h1['recType']) {
                case self::RT_DOCUMENTATOM:
                    $this->readDocumentAtom();
                    break;
                case self::RT_SLIDELISTWITHTEXT:
                    $this->readSlideList();
                    break;
            }
            $b -= (8+$h1['recLen']);
        }
    }

    private function readDocumentStream()
    {
        $userEdit = $this->readUserEditAtom($this->offsetToCurrentEdit);
        $docPersistID = $userEdit['docPersistIDRef'];
        $atoms = array();
        do {
            array_unshift($atoms, $this->readPersistDirectoryAtom($userEdit['offsetPersistDirectory']));
                        
            if ($userEdit['offsetLastEdit'] != 0) {
                $userEdit = $this->readUserEditAtom($userEdit['offsetLastEdit']);
            }
        } while ($userEdit['offsetLastEdit'] != 0);
    
        $this->persistDirectory = array();
        foreach ($atoms as $a) {
            // can't use array_merge for this because it renumbers numeric keys
            foreach ($a as $k => $v) {
                $this->persistDirectory[$k] = $v;
            }
        }
        
        $this->readDocumentContainer($this->persistDirectory[$docPersistID]);
    }

    public function getText()
    {
        return $this->text;
    }
}
