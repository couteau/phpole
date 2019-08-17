<?php
namespace Cryptodira\PhpMsOle;

/**
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
class MsDocDocument
{
    /* @var \Cryptodira\PhpOffice\OLEDocument */
    private $OLEDocument;
    private $streamid;
    private $summarystreamid;
    private $docinfostreamid;
    private $fib;
    private $summaryinfo;
    private $docsummaryinfo;

    const FIBFormat =
    'v1wIdent/' .    # 00
    'v1nFib/' .      # 02
    'v1nProduct/' .  # 04
    'v1lid/' .       # 06
    'v1pnNext/' .    # 08
    'v1flags1/' .    # 0A
    'v1nFibBack/' .  # 0C
    'V1lKey/' .      # 0E
    'C1envr/' .      # 12
    'C1flags2/' .    # 13
    'v1Chs/' .       # 14
    'v1chsTables/' . # 16
    'V1fcMin/' .     # 18
    'V1fcMax/' .     # 1C
    'v1csw/' .                  # 20
    'v1wMagicCreated/' .        # 22 fibRgW - next 14 words
    'v1wMagicRevised/' .        # 24
    'v1wMagicCreatedPrivate/' . # 26
    'v1wMagicRevisedPrivate/' . # 28
    'v1pnFbpChpFirst_W6/' .     # 2A these are signed values, but unpack has no format code for little-endian signed shorts
    'v1pnChpFirst_W6/' .        # 2C
    'v1cpnBteChp_W6/' .         # 2E
    'v1pnFbpPapFirst_W6/' .     # 30
    'v1pnPapFirst_W6/' .        # 32
    'v1cpnBtePap_W6/' .         # 34
    'v1pnFbpLvcFirst_W6/' .     # 36
    'v1pnLvcFirst_W6/' .        # 38
    'v1cpnBteLvc_W6/' .         # 3A
    'v1lidFE/' .                # 3C
    'v1clw/' .                  # 3E
    'V1cbMac/' .                # 40
    'V1lProductCreated/' .      # 44
    'V1lProductRevised/' .      # 48
    'V1ccpText/' .              # 4C
    'V1ccpFtn/' .               # 50
    'V1ccpHdd/' .               # 54
    'V1ccpMcr/' .               # 58
    'V1ccpAtn/' .               # 5C
    'V1ccpEdn/' .               # 60
    'V1ccpTxbx/' .              # 64
    'V1ccpHdrTxbx/' .           # 68
    'V1pnFbpChpFirst/' .        # 6C
    'V1pnChpFirst/' .           # 70
    'V1cpnBteChp/' .            # 74
    'V1pnFbpPapFirst/' .        # 78
    'V1pnPapFirst/' .           # 7C
    'V1cpnBtePap/' .            # 80
    'V1pnFbpLvcFirst/' .        # 84
    'V1pnLvcFirst/' .           # 88
    'V1cpnBteLvc/' .            # 8C
    'V1fcIslandFirst/' .        # 90
    'V1fcIslandLim/' .          # 94
    'v1Cfclcb';                 # 98

    const FibRgFcLcbSize = [
        0x00C1 => 0x005D,
        0x00D9 => 0x006C,
        0x0101 => 0x0088,
        0x010C => 0x00A4,
        0x0112 => 0x00B7,
    ];


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

        $this->streamid = $this->OLEDocument->GetDocumentStream();
        $this->summarystreamid = $this->OLEDocument->FindStreamByName(chr(5) . 'SummaryInformation');
        $this->docinfostreamid = $this->OLEDocument->FindStreamByName(chr(5) . 'DocumentSummaryInformation');
        $this->docsummaryinfo = null;
        $this->summaryinfo = null;

        $data = $this->OLEDocument->Read($this->streamid, 154);
        $this->fib = unpack(self::FIBFormat, $data);
    }


    public function GetText()
    {
        return $this->OLEDocument->Read($this->streamid, $this->fib['fcMax'], $this->fib['fcMin']);
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
}

