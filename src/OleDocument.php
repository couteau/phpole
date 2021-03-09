<?php
namespace Cryptodira\PhpOle;

/**
 *
 * OleDocument represents the underlying file structure of a Microsoft Ole compound file,
 * currently limited to readonly
 *
 * Credit to the PhpOffice/* projects by Maarten Balliauw for ideas and insights
 *
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
class OleDocument implements \IteratorAggregate, \Countable, \ArrayAccess
{

    /**
     * Directory and file name containing the data store
     *
     * @var string
     */
    private $filepath;

    /**
     * File descriptor for underlying data store
     *
     * @var resource
     */
    private $stream;

    /**
     * Ole header structure
     * An array with entries corresponding to the OleDocument::OleHeaderFormat format string
     *
     * @var array
     */
    private $header;

    /**
     * Whether the header fields have been modified and need to be written to the file
     *
     * @var bool
     */
    private $headerDirty;

    /**
     * Sector size for this file: 512 for version <= 3, 4096 for version 4
     * calculated as 1 << SectorShift
     *
     * @var int
     */
    private $sectorsize;

    /**
     * Sector size for the mini stream
     *
     * @var int
     */
    private $minisectorsize;

    /**
     * Format string of regular FAT sectors for unpack()
     * Either OleV3FATSectorFormat or OleV4FATSectorFormat
     *
     * @var string
     */
    private $FAT_sectorformat;

    /**
     * Format string of double indirect FAT sectors for unpack()
     * Either OleV3DIFATSectorFormat or OleV4DIFATSectorFormat
     *
     * @var string
     */
    private $DIFAT_sectorformat;

    /**
     * Ole double indirect file allocation table for the file
     *
     * @var array
     */
    private $DIFAT;

    /**
     * Ole file allocation table for the file
     *
     * @var array
     */
    private $FAT;

    /**
     * Ole file allocation table for the mini-stream
     *
     * @var array
     */
    private $miniFAT;

    /**
     * All directory entries in the Ole container, not in a directory structure
     *
     * @var array
     */
    private $fileSpecs;

    /**
     * The document's root storage directory
     *
     * @var OleStorage
     */
    private $rootStorage;

    // Special FAT entry values
    const MAXREGSECT = 0xFFFFFFFA;

    const DIFSECT = 0xFFFFFFFC;

    const FATSECT = 0xFFFFFFFD;

    const ENDOFCHAIN = 0xFFFFFFFE;

    const FREESECT = 0xFFFFFFFF;

    // Special stream ID values
    const MAXREGSID = 0xFFFFFFFA;

    const NOSTREAM = 0xFFFFFFFF;

    // Directory entry types
    const UnkownObject = 0x00;

    const StorageObject = 0x01;

    const StreamObject = 0x02;

    const RootStorageObject = 0x05;

    const TYPE_MAP = [
        self::StorageObject => OleStorage::class,
        self::StreamObject => OleStream::class,
        self::RootStorageObject => OleStorage::class,
    ];

    // Size of the Ole header record
    const OleHeaderSize = 0x200;

    // Size of a single directory entry
    const OleDirectoryEntrySize = 128;

    // Magic number identifying file as an Ole compound file
    const OleSignature = 'd0cf11e0a1b11ae1';

    // Unpack() format string for a 512 byte FAT sector
    // Each FAT sector contains 128 FAT entries
    const OleV3FATSectorFormat = 'V128';

    // Unpack() format string for a 512 byte directory information FAT sector
    // Each sector contains 127 FAT entries for sectors containing directory information and
    // one entry pointing to the next directory information FAT sector
    const OleV3DIFATSectorFormat = 'V127FATSectors/V1NextDIFATSector';

    // Unpack() format string for a 4096 byte FAT sector
    // Each FAT sector contains 1024 FAT entries
    const OleV4FATSectorFormat = 'V1024';

    // Unpack() format string for a 4096 byte directory information FAT sector
    // Each sector contains 1023 FAT entries for sectors containing directory information and
    // one entry pointing to the next directory information FAT sector
    const OleV4DIFATSectorFormat = 'V1023FATSectors/V1NextDIFATSector';

    // @formatter:off
    // Unpack() format string for the Ole Header
    const OleHeaderFormat =
            'H16Signature/' .           # 00 8 bytes
            'H32CLSID/' .               # 08 16 bytes
            'v1MinorVersion/' .         # 18 2 bytes
            'v1MajorVersion/' .         # 1A 2 bytes
            'v1ByteOrder/' .            # 1C 2 bytes
            'v1SectorShift/' .           # 1E 2 bytes
            'v1MiniSectorShift/' .       # 20 2 bytes
            'Z6Reserved1/' .            # 22 6 bytes
            'V1DirectorySectors/' .      # 28 4 bytes
            'V1FATSectors/' .            # 2C 4 bytes
            'V1FirstDirectorySector/' .  # 30 4 bytes
            'V1TransactionSignature/' . # 34 4 bytes
            'V1MiniStreamCutoff/' .     # 38 4 bytes
            'V1FirstMiniFATSector/' .    # 3C 4 bytes
            'V1MiniFATSectorCount/' .    # 40 4 bytes
            'V1FirstDIFATSector/' .      # 44 4 bytes
            'V1DIFATSectorCount/' .      # 48 4 bytes
            'V109DIFAT';                # 4C 436 bytes

    // Unpack() format string for a single directory entry (128 bytes)
    const OleDirectoryEntryFormat =
            'A64EntryName/' .       # 00 64 bytes
            'v1EntryNameLength/' .  # 40 2 bytes
            'C1ObjectType/' .       # 42 1 bytes
            'C1ColorFlag/' .        # 43 1 byte
            'V1LeftSiblingID/' .    # 44 4 bytes
            'V1RightSiblingID/' .   # 48 4 bytes
            'V1ChildID/' .          # 4C 4 bytes
            'H32CLSID/' .           # 50 16 bytes
            'V1StateBits/' .        # 60 4 bytes
            'P1CreationTime/' .     # 64 8 bytes
            'P1ModifiedTime/' .     # 6C 8 bytes
            'V1StartingSector/' .    # 74 4 bytes
            'P1StreamSize/';        # 78 8 bytes
    // @formatter:on

    /**
     * Initialize a new OleDocument structure
     */
    public function __construct($source = null)
    {
        $this->stream = null;
        if ($source !== null) {
            $this->open($source);
        }
    }

    /**
     * Dispose of internal resources
     */
    public function __destruct()
    {
        $this->close();
    }

    public function new($filepath, $format = true)
    {
        $this->filepath = $filepath;
        if ($format) {
            $this->format();
        }
        return $this;
    }

    public function open($source)
    {
        if (is_resource($source) && get_resource_type($source) === 'stream') {
            $this->CreateFromStream($source);
        } elseif (is_string($source)) {
            if (strlen($source) >= 8 && unpack('H16', $source)[1] === self::OleSignature) {
                $this->createFromString($source);
            } elseif (is_readable($source)) {
                $this->CreateFromFile($source);
            } else {
                throw new \InvalidArgumentException(
                        'Passed source was not a valid file name and did not contain an Ole stream');
            }
        } else {
            throw new \InvalidArgumentException('Passed source could not be interpreted as an Ole file');
        }
        return $this;
    }

    private function writeHeader($version = null)
    {
        if ($version === null && $this->header) {
            $version = $this->header['MajorVersion'];
        }

        if ($version != 0x0003 && $version != 0x0004) {
            throw new \InvalidArgumentException('Invalid OLE file version: ' . $version);
        }

        if (!$this->header) {
            $$this->header = $this->initializeHeader($version);
        }
        $headerValues = array_merge(array_values($this->header), array_slice($this->DIFAT, 0, 109));

        $fmt = preg_replace('/([A-Za-z][0-9]+)[A-Z]\w+(\/|$)/', '$1', self::OleHeaderFormat);
        $header = pack($fmt, ...$headerValues);

        if (!$this->stream) {
            $this->stream = fopen($this->filepath, 'r+');
        } else {
            rewind($this->stream);
        }

        fwrite($this->stream, $header);
        if ($version === 0x0004) {
            fwrite($this->stream, str_repeat(chr(0), 0x800)); // pad header to a full sector
        }
    }

    private function writeFAT()
    {
        for ($fat = $this->FAT, $i = 0; $fatdata = array_splice($fat, 0, $this->sectorsize / 4); $i++) {
            $fatstring = pack($this->FAT_sectorformat, ...$fatdata);
            if ($this->DIFAT[$i] >= self::MAXREGSECT) {
                throw new \BadMethodCallException('Number of FAT entries does not correspond to number of FAT sectors in DIFAT');
            }
            $this->writeSectorData($this->DIFAT[$i], $fatstring);
        }
    }

    private function writeDirectorySectors()
    {
        $fmt = preg_replace('/([A-Za-z][0-9]+)[A-Z]\w+(\/|$)/', '$1', self::OleDirectoryEntryFormat);

        $s = $this->header['FirstDirectorySector'];
        if ($s === self::ENDOFCHAIN) {
            return;
        }

        $data = '';
        foreach ($this->fileSpecs as $index => $fileSpec) {
            $fileSpec['EntryName'] = mb_convert_encoding($fileSpec['EntryName'], "UTF-16LE", "UTF-8");
            $newentry = pack($fmt, ...array_values($fileSpec));
            $data .= $newentry;
            if (strlen($data) === $this->sectorsize) {
                $this->writeSectorData($s, $data);
                $s = $this->FAT[$s];
                if ($s === self::ENDOFCHAIN && $index < count($this->fileSpecs) - 1) {
                    throw new \BadMethodCallException('Insufficient directory sectors allocated to hold all directory entries');
                }
                $data = '';
            }
        }

        if (strlen($data) < $this->sectorsize) {
            $this->writeSectorData($s, $data);
        }
    }

    public function writeMiniStream()
    {
        // write the miniFAT data. Don't need to write the stream for now, because it's unbuffered and will be written on modification
        for ($i = 0, $s = $this->header['FirstMiniFATSector'], $fat = $this->miniFAT; $fatdata = array_splice($fat, 0, $this->sectorsize / 4); $i++, $s = $this->FAT[$s]) {
            $fatstring = pack($this->FAT_sectorformat, ...$fatdata);
            if ($s >= self::MAXREGSECT) {
                throw new \BadMethodCallException('Number of miniFAT entries does not correspond to number of miniFAT sectors in the FAT');
            }
            $this->writeSectorData($s, $fatstring);
        }
    }

    public function save()
    {
        $this->writeHeader();
        $this->writeFAT();
        $this->writeMiniStream();
        $this->writeDirectorySectors();
    }

    /**
     * Close the underlying file resource and reset the internal structures
     */
    public function close()
    {
        if ($this->stream)
            fclose($this->stream);
        $this->stream = null;
        $this->header = null;
        $this->FAT = null;
        $this->miniFAT = null;
        $this->rootStorage = null;
        $this->fileSpecs = null;
        return $this;
    }

    private function initializeHeader($version)
    {
        $header = array_merge([
            self::OleSignature,	                    // Signature
            str_repeat('0', 32),	                // CLSID
            0x003E,	                                // MinorVersion
            $version,	                            // MajorVersion
            0xFFFE,	                                // ByteOrder
            $version === 0x0003 ? 0x0009 : 0x000C,	// SectorShift
            0x0006,	                                // MiniSectorShift
            str_repeat(chr(0), 6),	                // Reserved1
            0,	                                    // DirectorySectors
            0,	                                    // FATSectors
            self::ENDOFCHAIN,	                    // FirstDirectorySector
            0,	                                    // TransactionSignature
            0x00001000,	                            // MiniStreamCutoff
            self::ENDOFCHAIN,	                    // FirstMiniFATSector
            0,	                                    // MiniFATSectorCount
            self::ENDOFCHAIN,	                    // FirstDIFATSector
            0,	                                    // DIFATSectorCount
        ], array_fill(0, 109, self::FREESECT));     // DIFAT
        
        $fmt = preg_replace('/([A-Za-z][0-9]+)[A-Z]\w+(\/|$)/', '$1', self::OleHeaderFormat);
        $data = pack($fmt, ...$header);
        $this->readHeader($data);
        $this->headerDirty = true;
    }

    private function intitializeFAT()
    {
        // create one sector worth of fat entries, and reserve sector 0 for the new FAT sector
        $this->FAT = array_fill(0, $this->sectorsize / 4, self::FREESECT);
        $this->FAT[0] = self::FATSECT;
        $this->header['FATSectors'] = 1;
        $this->DIFAT[0] = 0;
        $this->headerDirty = true;
    }

    public function initializeRootStorage()
    {
        $this->header['FirstDirectorySector'] = $this->allocateSector();
        if ($this->header['MajorVersion'] === 0x0004) {
            $this->header['DirectorySectors'] = 1;
        }
        $this->fileSpecs = array([
            'EntryName' => 'RootEntry',
            'EntryNameLength' => 9,
            'ObjectType' => self::RootStorageObject,
            'ColorFlag' => 0,
            'LeftSiblingID' => self::NOSTREAM,
            'RightSiblingID' => self::NOSTREAM,
            'ChildID' => self::NOSTREAM,
            'CLSID' => str_repeat('0', 32),
            'StateBits' => 0,
            'CreationTime' => 0,
            'ModifiedTime' => 0,
            'StartingSector' => self::ENDOFCHAIN,
            'StreamSize' => 0,
        ]);
        $this->rootStorage = new OleStorage($this);
        $this->headerDirty = true;
    }

    public function initializeMiniStream()
    {
        // TODO: make sure the ministream isn't already initialized
        $this->header['FirstMiniFATSector'] = $this->allocateSector();
        $this->header['MiniFATSectorCount'] = 1;
        $this->miniFAT = array_fill(0, $this->sectorsize / 4, self::FREESECT);
        $this->fileSpecs[0]['StartingSector'] = $this->allocateSector(true);
        $this->headerDirty = true;
    }

    public function format($version = 0x0003)
    {
        $this->header = null;
        $this->FAT = null;
        $this->miniFAT = null;
        $this->rootStorage = null;
        $this->fileSpecs = null;
        if ($this->stream) {
            ftruncate($$this->stream, 0);
        } else {
            $this->stream = fopen($this->filepath, 'w+');
        }
        $this->initializeHeader($version);
        $this->intitializeFAT();
        $this->initializeRootStorage();
        $this->save();
    }

    private function allocateSector($initialize = false, $code = self::ENDOFCHAIN, $prevSector = null)
    {
        for ($i = 0; $i < count($this->FAT) && $this->FAT[$i] !== self::FREESECT; $i++);
        if ($i ===  count($this->FAT)) {
            // Allocate a new FAT sector
        }
        $this->FAT[$i] = $code;
        if ($prevSector) {
            $this->FAT[$prevSector] = $i;
        }
        if ($initialize) {
            $this->writeSectorData($i, '');
        }
        return $i;
    }

    /**
     * Get the size of big sectors in this Ole file (512 or 4096)
     */
    public function getSectorSize()
    {
        return $this->sectorsize;
    }

    /**
     * Get the the sector number of the next sector in the chain
     *
     * @param int $sector
     */
    public function getNextSector($sector)
    {
        if ($sector <= self::MAXREGSECT) {
            return $this->FAT[$sector];
        } else {
            return $sector;
        }
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
        fseek($this->stream, ($sector + 1) * $this->sectorsize);
        $data = fread($this->stream, $this->sectorsize);

        if (!$data)
            throw new \Exception("Could not read sector {$sector}");

        return $data;
    }

    private function writeSectorData($sector, $data)
    {
        if (strlen($data) > $this->sectorsize) {
            throw new \BadMethodCallException('attempt to write more than ' . $this->sectorsize . ' bytes to a single sector');
        }
        
        if (strlen($data) < $this->sectorsize) {
            $data = str_pad($data, $this->sectorsize, chr(0));
        }
        fseek($this->stream, (1 + $sector) * $this->sectorsize);
        fwrite($this->stream, $data);
    }

    /**
     * Read a seingle sector from the mini stream
     *
     * @param int $sector
     * @return string
     */
    public function getMiniSectorData($sector)
    {
        return $this->read(0, 64, $sector * 64);
    }

    /**
     * Read the header record and initialize internal data structures/values
     */
    private function readHeader($data = null)
    {
        if (!$data) {
            $data = (string) fread($this->stream, self::OleHeaderSize);
        }

        if (!$data) {
            throw new \Exception("Could not read header sector");
        }

        $this->header = unpack($this::OleHeaderFormat, $data);
        if ($this->header['Signature'] != self::OleSignature) {
            throw new \Exception("Stream is not an Ole file");
        }

        $this->DIFAT = array_values(array_splice($this->header, -109, 109));
        $this->sectorsize = 1 << $this->header['SectorShift'];
        $this->minisectorsize = 1 << $this->header['MiniSectorShift'];
        switch ($this->header['MajorVersion']) {
            case 0x03:
                $this->FAT_sectorformat = self::OleV3FATSectorFormat;
                $this->DIFAT_sectorformat = self::OleV3DIFATSectorFormat;
                break;
            case 0x04:
                $this->FAT_sectorformat = self::OleV4FATSectorFormat;
                $this->DIFAT_sectorformat = self::OleV4DIFATSectorFormat;
                fseek($this->stream, $this->sectorsize, SEEK_SET); // move to start of first sector after header
                break;
            default:
                throw new \Exception("Invalid SectorShift");
        }
    }

    /** 
     * Load the DIFAT sectors beyond the DIFAT entries stored in the header
     * and add them to the DIFAT array
     */
    private function readDIFAT()
    {
        $s = $this->header['FirstDIFATSector'];
        while ($s != self::ENDOFCHAIN) {
            $data = $this->getSectorData($s);
            $difat = unpack($this->DIFAT_sectorformat, $data);
            $this->DIFAT = array_merge($this->DIFAT, array_values(array_slice($difat, 0, sizeof($difat) - 1)));
            $s = $difat['NextDIFATSector'];
        }
    }

    /**
     * Load the entire FAT into $this->FAT
     */
    private function readFAT()
    {
        $this->FAT = array();
        for ($i = 0; $i < count($this->DIFAT) && $this->DIFAT[$i] !== self::FREESECT; $i++) {
            $data = $this->getSectorData($this->DIFAT[$i]);
            $entries = unpack($this->FAT_sectorformat, $data);
            $this->FAT = array_merge($this->FAT, $entries);
        }
    }

    /**
     * Load the entire miniFAT into $this->miniFAT
     */
    private function readMiniFAT()
    {
        $this->miniFAT = array();
        $s = $this->header['FirstMiniFATSector'];
        while ($s != self::ENDOFCHAIN) {
            $data = $this->getSectorData($s);
            $entries = unpack($this->FAT_sectorformat, $data);
            $this->miniFAT = array_merge($this->miniFAT, $entries);
            $s = $this->FAT[$s];
        }
    }

    /**
     * Read directory entries from a directory information sector and add them to an existing
     * array, if passed.
     * Return an array containing the directory entries.
     *
     * @param int $sector
     * @param array $entries
     * @return string|string[]
     */
    private function readDirectorySector($sector)
    {
        $data = $this->getSectorData($sector);
        for ($i = 0; $i < $this->sectorsize / self::OleDirectoryEntrySize; $i++) {
            $newentry = unpack(OleDocument::OleDirectoryEntryFormat, $data, $i * OleDocument::OleDirectoryEntrySize);
            if ($newentry['ObjectType'] === 0) {
                break;
            }
            // unpack cuts off the final byte of the final UTF-16LE character if it is null, so we have to add it back on before converting
            if (strlen($newentry['EntryName']) % 2 != 0)
                $newentry['EntryName'] .= chr(0);
            $newentry['EntryName'] = mb_convert_encoding($newentry['EntryName'], "UTF-8", "UTF-16LE");
            $this->fileSpecs[] = $newentry;
        }
    }

    /**
     * Read the all filespec entries for the Ole file
     */
    private function readFileSpecs()
    {
        $this->fileSpecs = [];
        $s = $this->header['FirstDirectorySector'];
        while ($s != OleDocument::ENDOFCHAIN) {
            $this->readDirectorySector($s);
            $s = $this->FAT[$s];
        }
    }

    /**
     * Read the root director for the Ole file
     */
    private function readRootDirectory()
    {
        $this->rootStorage = new OleStorage($this, 0);
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
    public function read($streamid, $bytes = -1, $offset = 0)
    {
        if ($streamid < 0 || $streamid >= sizeof($this->fileSpecs))
            throw new \Exception("Invalid StreamID $streamid");

        if ($this->fileSpecs[$streamid]['ObjectType'] != 2 && $this->fileSpecs[$streamid]['ObjectType'] != 5)
            throw new \Exception("StreamID $streamid is not a stream");

        if ($offset > $this->fileSpecs[$streamid]['StreamSize'])
            throw new \Exception("Attempt to read past end of stream");

        // $bytes = -1 means read the whole stream
        if ($bytes == -1)
            $bytes = $this->fileSpecs[$streamid]['StreamSize'];

        if ($streamid == 0 || $this->fileSpecs[$streamid]['StreamSize'] >= $this->header['MiniStreamCutoff']) {
            $readsector = array(
                $this,
                'getSectorData'
            );
            $fat = $this->FAT;
            $bs = $this->sectorsize;
        } else {
            $readsector = array(
                $this,
                'getMiniSectorData'
            );
            $fat = $this->miniFAT;
            $bs = 64;
        }

        // first find the sector containing the starting offset
        $s = $this->fileSpecs[$streamid]['StartingSector'];
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
            return substr($data, 0, $bytes);
        else
            return $data;
    }

    /**
     * Read all of the data for a stream within the Ole file
     *
     * @param int $streamid
     * @return string
     */
    public function getStreamData($streamid)
    {
        if ($this->fileSpecs[$streamid]['ObjectType'] != 2)
            return null; // should this throw an error?

        if ($streamid == 0 || $this->fileSpecs[$streamid]['StreamSize'] >= $this->header['MiniStreamCutoff']) {
            $readsector = array(
                $this,
                'getSectorData'
            );
            $fat = $this->FAT;
            // $bs = $this->sectorsize;
        } else {
            $readsector = array(
                $this,
                'getMiniSectorData'
            );
            $fat = $this->miniFAT;
            // $bs = 64;
        }

        $s = $this->fileSpecs[$streamid]['StartingSector'];
        $data = '';
        while ($s != self::ENDOFCHAIN) {
            $data .= $readsector($s);
            $s = $fat[$s];
        }

        return substr($data, 0, $this->fileSpecs[$streamid]['StreamSize']);
    }

    /**
     * Return the streamid for the main document stream in this Ole file or false if none could be found
     *
     * @return number|boolean
     */
    public function getDocumentStream()
    {
        foreach ($this->fileSpecs as $id => $stream) {
            if ($stream['ObjectType'] == 2 && $stream['StartingSector'] == 0)
                return $id;
        }

        return false;
    }

    public function getObject($entry)
    {
        if (is_null($entry)) {
            $entry = $this->getDocumentStream();
            $filespec = $this->fileSpecs[$entry];
        } elseif (is_string($entry)) {
            if (!$entry = $this->findEntryByName($entry)) {
                throw new \Exception("Stream {$entry} not found");
            }
            $filespec = $this->fileSpecs[$entry];
        } elseif (is_int($entry)) {
            $filespec = $this->fileSpecs[$entry];
        } elseif (is_array($entry)) {
            $filespec = $entry;
            $entry = array_search($filespec, $this->fileSpecs);
        } else {
            throw new \Exception("Invalid stream {$entry}");
        }

        if ($filespec['ObjectType'] === self::RootStorageObject) {
            return $this->rootStorage;
        } else {
            $class = self::TYPE_MAP[$filespec['ObjectType']];
            return new $class($this, $entry);
        }
    }

    /**
     * Return the streamid for the stream with the passed stream name
     *
     * @param string $entryName
     * @return number|boolean
     */
    public function findEntryByName($entryName)
    {
        // could use the separate name-indexed array to do this more quickly
        foreach ($this->fileSpecs as $id => $entry) {
            if ($entry['EntryName'] == $entryName) {
                return $id;
            }
        }

        return false;
    }

    /**
     * Return the root directory of the Ole document
     *
     * @return int
     */
    public function getRootStorage()
    {
        return $this->rootStorage;
    }

    /**
     * Return the number of entries in the root directory
     *
     * @return int
     */
    public function getRootStorageCount()
    {
        return count($this->rootStorage);
    }

    /**
     * Create an OleDocument on top of the passed stream, initializing the internal structures
     *
     * @param resource $strm
     * @throws \Exception
     * @return \Cryptodira\PhpMsOle\OleDocument
     */
    public function CreateFromStream($strm)
    {
        if (!$strm || !is_resource($strm) || get_resource_type($strm) !== 'stream') {
            throw new \Exception("Invalid stream passed to OleDocument::create");
        }

        $this->stream = $strm;
        rewind($this->stream);

        $this->readHeader();
        $this->readDIFAT();
        $this->readFAT();
        $this->readMiniFAT();
        $this->readFileSpecs();
        $this->readRootDirectory();
        return $this;
    }

    /**
     * Open the passed filename and create an OleDocument on top of the contents
     *
     * @param string $filepath
     * @throws \Exception
     * @return \Cryptodira\PhpMsOle\OleDocument
     */
    public function CreateFromFile($filepath)
    {
        if (!is_readable($filepath)) {
            throw new \Exception("Could not open {$filepath} for reading");
        }

        $this->filepath = $filepath;
        $strm = fopen($filepath, 'rb');
        return $this->CreateFromStream($strm);
    }

    /**
     * Create an OleDocument from the passed data string
     *
     * @param string $fdata
     * @return \Cryptodira\PhpMsOle\OleDocument
     */
    public function CreateFromString($fdata)
    {
        $strm = fopen('php://temp,' . $fdata, 'r');
        return $this->CreateFromStream($strm);
    }

    public function getIterator()
    {
        return new \ArrayIterator($this->fileSpecs);
    }

    public function count()
    {
        return count($this->fileSpecs);
    }

    public function offsetGet($offset)
    {
        return $this->fileSpecs[$offset];
    }

    public function offsetExists($offset)
    {
        return array_key_exists($offset, $this->fileSpecs);
    }

    public function offsetUnset($offset)
    {
        throw new \BadMethodCallException('OleDocument file specs are readonly');
    }

    public function offsetSet($offset, $value)
    {
        throw new \BadMethodCallException('OleDocument file specs are readonly');
    }
}

