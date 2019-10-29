<?php
namespace Cryptodira\PhpMsOle;

/**
 *
 * OLEDocument represents the underlying file structure of a Microsoft OLE compound file,
 * currently limited to readonly
 *
 * Credit to the PhpOffice/* projects by Maarten Balliauw for ideas and insights
 *
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
class OLEDocument
{
    /**
     * File descriptor for underlying data store
     *
     * @var resource
     */
    private $stream;

    /**
     * OLE header structure
     * An array with entries corresponding to the OLEDocument::OLEHeaderFormat format string
     *
     * @var array
     */
    private $header;

    /**
     * Sector size for this file: 512 for version <= 3, 4096 for version 4
     *
     * @var int
     */
    private $blocksize;

    /**
     * Format string of regular sectors for unpack()
     * Either OLEV3FATSectorFormat or OLEV4FATSectorFormat
     *
     * @var string
     */
    private $FAT_sectorformat;

    /**
     * Format string of directory information sectors for unpack()
     * Either OLEV3DIFATSectorFormat or OLEV4DIFATSectorFormat
     *
     * @var string
     */
    private $DIFAT_sectorformat;

    /**
     * OLE file allocation table for the file
     *
     * @var array
     */
    private $FAT;


    /**
     * OLE file allocation table for the mini-stream
     *
     * @var array
     */
    private $MiniFAT;

    /**
     * Array of directory entries in the root directory structure
     * Each entry is an array with entries corresponding to the OLEDocument::OLEDirectoryEntryFormat format string
     *
     * @var array
     */
    private $RootDir;

    // Special FAT entry values
    const MAXREGSECT    = 0xFFFFFFFA;
    const DIFSECT       = 0xFFFFFFFC;
    const FATSECT       = 0xFFFFFFFD;
    const ENDOFCHAIN    = 0xFFFFFFFE;
    const FREESECT      = 0xFFFFFFFF;

    // Special stream ID values
    const MAXREGSID     = 0xFFFFFFFA;
    const NOSTREAM      = 0xFFFFFFFF;

    // Directory entry types
    const UnkownObject	     = 0x00;
    const StorageObject	     = 0x01;
    const StreamObject	     = 0x02;
    const RootStorageObject	 = 0x05;

    // Size of the OLE header record
    const OLEHeaderSize = 0x200;

    // Size of a single directory entry
    const OLEDirectoryEntrySize = 128;

    // Magic number identifying file as an OLE compound file
    const OLESignature = 'd0cf11e0a1b11ae1';

    // Unpack() format string for a 512 byte FAT sector
    // Each FAT sector contains 128 FAT entries
    const OLEV3FATSectorFormat = 'V128';

    // Unpack() format string for a 512 byte directory information FAT sector
    // Each sector contains 127 FAT entries for sectors containing directory information and
    // one entry pointing to the next directory information FAT sector
    const OLEV3DIFATSectorFormat = 'V127FATSectors/V1NextDIFATSector';

    // Unpack() format string for a 4096 byte FAT sector
    // Each FAT sector contains 1024 FAT entries
    const OLEV4FATSectorFormat = 'V1024';

    // Unpack() format string for a 4096 byte directory information FAT sector
    // Each sector contains 1023 FAT entries for sectors containing directory information and
    // one entry pointing to the next directory information FAT sector
    const OLEV4DIFATSectorFormat = 'V1023FATSectors/V1NextDIFATSector';

    // Unpack() format string for the OLE Header
    const OLEHeaderFormat =
    'H16Signature/' .            # 00 8 bytes
    'H32CLSID/' .                # 08 16 bytes
    'v1MinorVersion/' .          # 18 2 bytes
    'v1MajorVersion/' .          # 1A 2 bytes
    'v1ByteOrder/' .             # 1C 2 bytes
    'v1SectorShift/' .           # 1E 2 bytes
    'v1MiniSectorShift/' .       # 20 2 bytes
    'Z6Reserved1/' .             # 22 6 bytes
    'V1DirectorySectors/' .      # 28 4 bytes
    'V1FATSectors/' .            # 2C 4 bytes
    'V1FirstDirectorySector/' .  # 30 4 bytes
    'V1TransactionSignature/' .  # 34 4 bytes
    'V1MiniStreamCutoff/' .      # 38 4 bytes
    'V1FirstMiniFATSector/' .    # 3C 4 bytes
    'V1MiniFATSectorCount/' .    # 40 4 bytes
    'V1FirstDIFATSector/' .      # 44 4 bytes
    'V1DIFATSectorCount/' .      # 48 4 bytes
    'V109DIFAT';                 # 4C 436 bytes

    // Unpack() format string for a single directory entry
    const OLEDirectoryEntryFormat =
    'A64EntryName/' .
    'v1EntryNameLength/' .
    'C1ObjectType/' .
    'C1ColorFlag/' .
    'V1LeftSiblingID/' .
    'V1RightSiblingID/' .
    'V1ChildID/' .
    'H32CLSID/' .
    'V1StateBits/' .
    'P1CreationTime/' .
    'P1ModifiedTime/' .
    'V1StartingSector/' .
    'P1StreamSize/';

    /**
     *  Initialize a new OLEDocument structure
     *
     */
    public function __construct()
    {
        $this->stream = null;
    }


    /**
     *  Dispose of internal resources
     *
     */
    public function __destruct()
    {
        $this->Close();
    }

    /**
     * Close the underlying file resource and reset the internal structures
     *
     */
    public function Close()
    {
        if ($this->stream)
            fclose($this->stream);
        $this->stream = null;
    }

    /**
     * Read a single sector from the underlying file descriptor
     *
     * @param int $sector
     * @throws \Exception
     * @return string
     */
    public function getSectorData($sector)
    {
        fseek($this->stream, ($sector + 1) * $this->blocksize);
        $data = fread($this->stream, $this->blocksize);

        if (!$data)
            throw new \Exception("Could not read sector {$sector}");

        return $data;
    }

    /**
     * Read a seingle sector from the mini stream
     *
     * @param int $sector
     * @return string
     */
    public function getMiniSectorData($sector)
    {
        return $this->Read(0,64,$sector * 64);
    }

    /**
     *  Load the entire FAT into $this->FAT
     */
    private function ReadFAT()
    {
        $this->FAT = array();
        for ($i = 0; $i < 109; $i++) {
            if ($this->header['DIFAT'][$i] >= self::ENDOFCHAIN)
                return;
            $data = $this->getSectorData($this->header['DIFAT'][$i]);
            $entries = unpack($this->FAT_sectorformat, $data);
            $this->FAT = array_merge($this->FAT, $entries);
        }

        $s = $this->header['FirstDIFATSector'];
        while ($s != self::ENDOFCHAIN) {
            $data = $this->getSectorData($s);
            $difat = unpack($this->DIFAT_sectorformat, $data);
            $difat['FATSectors'] = array_values(array_splice($difat, 0, sizeof($difat)-1));

            for ($i = 0; $i < $this->blocksize/4 - 1; $i++) {
                if ($difat['FATSectors'][$i] == self::ENDOFCHAIN)
                    return;
                $data = $this->getSectorData($this->header['DIFAT'][$i]);
                $entries = unpack($this->FAT_sectorformat, $data);
                $this->FAT = array_merge($this->FAT, $entries);
            }

            $s = $difat['NextDIFATSector'];
        }
    }

    /**
     *  Load the entire MiniFAT into $this->MiniFAT
     */
    private function ReadMiniFAT()
    {
        $this->MiniFAT = array();
        $s = $this->header['FirstMiniFATSector'];
        while ($s != self::ENDOFCHAIN) {
            $data = $this->getSectorData($s);
            $entries = unpack($this->FAT_sectorformat, $data);
            $this->MiniFAT = array_merge($this->MiniFAT, $entries);
            $s = $this->FAT[$s];
        }
    }

    /**
     * Read directory entries from a directory information sector and add them to an existing
     * array, if passed. Return an array containing the directory entries.
     *
     * @param int $sector
     * @param array $entries
     * @return string|string[]
     */
    private function ReadDirectorySector($sector, &$entries = null)
    {
        $data = $this->getSectorData($sector);

        if (!$entries)
            $entries = array();

        for ($i = 0; $i < $this->blocksize/128; $i++) {
            $newentry = unpack(self::OLEDirectoryEntryFormat, $data, $i * self::OLEDirectoryEntrySize);
            // unpack cuts off the final byte of the final UTF-16LE character if it is null, so we have to add it back on
            if (strlen($newentry['EntryName']) % 2 != 0)
                $newentry['EntryName'] .= chr(0);
            $newentry['EntryName'] = mb_convert_encoding($newentry['EntryName'], "UTF-8", "UTF-16LE");
            $entries[] = $newentry;
        }
        return $entries;
    }

    /**
     * Read the root director for the OLE file
     */
    private function ReadRootDirectory()
    {
        $s = $this->header['FirstDirectorySector'];
        while ($s != self::ENDOFCHAIN) {
            $this->ReadDirectorySector($s, $this->RootDir);
            $s = $this->FAT[$s];
        }
    }

    /**
     * Read all of the data for a stream within the OLE file
     *
     * @param int $streamid
     * @return string
     */
    private function GetStream($streamid)
    {
        if ($this->RootDir[$streamid]['ObjectType'] != 2)
            return null; //should this throw an error?
        $s = $this->RootDir[$streamid]['StartingSector'];
        $data = '';
        while ($s != self::ENDOFCHAIN) {
            $data .= $this->getSectorData($s);
            $s = $this->FAT[$s];
        }

        return $data;
    }

    /**
     * Return the number of entries in the root directory
     *
     * @return int
     */
    public function GetRootDirCount()
    {
        return sizeof($this->RootDir);
    }

    /**
     * Return the streamid for the main document stream in this OLE file or false if none could be found
     *
     * @return number|boolean
     */
    public function GetDocumentStream()
    {
        for ($i = 1; $i < sizeof($this->RootDir); $i++) {
            if ($this->RootDir[$i]['ObjectType'] == 2 && $this->RootDir[$i]['StartingSector'] == 0)
                return $i;
        }

        return false;
    }

    /**
     * Return the streamid for the stream with the passed stream name
     *
     * @param string $stream_name
     * @return number|boolean
     */
    public function FindStreamByName($stream_name)
    {
        // could use the red/black tree to do this more quickly
        for ($i = 1; $i < sizeof($this->RootDir); $i++) {
            if ($this->RootDir[$i]['EntryName'] == $stream_name)
                return $i;
        }

        return false;
    }

    /**
     * Read $bytes bytes from the stream corresponding to streamid starting at $offset
     *
     * @param int $streamid
     * @param int $bytes
     * @param int $offset
     * @throws \Exception
     * @return string
     */
    public function Read($streamid, $bytes = -1, $offset = 0)
    {
        if ($streamid < 0 || $streamid >= sizeof($this->RootDir))
            throw new \Exception("Invalid StreamID $streamid");

        if ($this->RootDir[$streamid]['ObjectType'] != 2 && $this->RootDir[$streamid]['ObjectType'] = 5)
            throw new \Exception("StreamID $streamid is not a stream");

        if ($offset > $this->RootDir[$streamid]['StreamSize'])
            throw new \Exception("Attempt to read past end of stream");

        // $bytes = -1 means read the whole stream
        if ($bytes == -1)
            $bytes = $this->RootDir[$streamid]['StreamSize'];

        if ($streamid == 0 || $this->RootDir[$streamid]['StreamSize'] >= $this->header['MiniStreamCutoff']) {
            $readsector = array($this, 'getSectorData');
            $fat = $this->FAT;
            $bs = $this->blocksize;
        }
        else {
            $readsector = array($this, 'getMiniSectorData');
            $fat = $this->MiniFAT;
            $bs = 64;
        }

        // first find the sector containing the starting offset
        $s = $this->RootDir[$streamid]['StartingSector'];
        $i = 0;
        while ($s != self::ENDOFCHAIN && $i < $offset) {
            if ($offset < $i + $bs)
                break;
            $i += $bs;
            $s = $fat[$s];
        }

        // grab from the starting offset to the end of the starting sector
        $data = substr($readsector($s), $offset - $i, $bs - ($offset % $bs));


        // keep adding sectors until we've got all the bytes requested
        if ($s != self::ENDOFCHAIN) {
            $s = $fat[$s];
            while (strlen($data) < $bytes && $s != self::ENDOFCHAIN) {
                $data .= $readsector($s);
                $s = $fat[$s];
            }
        }

        // If the last sector read added more bytes than requested, truncate the returned data
        if (strlen($data) > $bytes)
            return substr($data,0,$bytes);
        else
            return $data;
    }

    /**
     * Create an OLEDocument on top of the passed stream, initializing the internal structures
     *
     * @param resource $strm
     * @throws \Exception
     * @return \Cryptodira\PhpMsOle\OLEDocument
     */
    public function CreateFromStream($strm)
    {
        if (!$strm || !is_resource($strm) || get_resource_type($strm) !== 'stream') {
            throw new \Exception("Invalid stream passed to OLEDocument::create");
        }

        $this->stream = $strm;
        rewind($this->stream);

        $data = (binary)fread($strm, self::OLEHeaderSize);
        if (!$data) {
            throw new \Exception("Could not read header block");
        }

        $this->header = unpack($this::OLEHeaderFormat, $data);
        if ($this->header['Signature'] != self::OLESignature) {
            throw new \Exception("Stream is not an OLE file");
        }

        $this->header['DIFAT'] = array_values(array_splice($this->header,-109,109));
        switch ($this->header['MajorVersion']) {
            case 0x03:
                $this->blocksize = 512;
                $this->FAT_sectorformat = self::OLEV3FATSectorFormat;
                $this->DIFAT_sectorformat = self::OLEV3DIFATSectorFormat;
                break;
            case 0x04:
                $this->blocksize = 4096;
                $this->FAT_sectorformat = self::OLEV4FATSectorFormat;
                $this->DIFAT_sectorformat = self::OLEV4DIFATSectorFormat;
                fseek($this->stream, 4096, SEEK_SET); // move to start of first block after header
                break;
            default:
                throw new \Exception("Invalid SectorShift");
        }

        $this->ReadFAT();
        $this->ReadMiniFAT();
        $this->ReadRootDirectory();
        return $this;
    }

    /**
     * Open the passed filename and create an OLEDocument on top of the contents
     *
     * @param string $filepath
     * @throws \Exception
     * @return \Cryptodira\PhpMsOle\OLEDocument
     */
    public function CreateFromFile($filepath)
    {
        if (!is_readable($filepath)) {
            throw new \Exception("Could not open {$filepath} for reading");
        }

        $strm = fopen($filepath, 'rb');
        return $this->CreateFromStream($strm);
    }

    /**
     * Create an OLEDocument from the passed data string
     *
     * @param string $fdata
     * @return \Cryptodira\PhpMsOle\OLEDocument
     */
    public function CreateFromString($fdata)
    {
        $strm = fopen('php://temp,' . $fdata);
        return $this->CreateFromStream($strm);
    }
}

