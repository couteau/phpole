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
     * @var OleHeader
     */
    private $header;

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
     * Maximum size of files that are stored in the mini stream
     *
     * @var int
     */
    private $miniStreamCutoff;

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
     * Sectors containing DIFAT entries
     *
     * @var array
     */
    private $DIFATSectors;

    /**
     * Ole file allocation table for the file
     *
     * @var array
     */
    private $FAT;

    /**
     * Whether the FAT has been modified and needs to be written to the file
     *
     * @var bool
     */
    private $FATDirty;

    /**
     * List of sectors whose data has not been written to file
     *
     * @var string[]
     */
    private $dirtySectors = [];

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
    private $directoryEntries;

    /**
     * The directory entry object of the root directory entry
     * 
     * @var OleDirectoryEntry
     */
    private $rootEntry;

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

    // R-B Tree color codes
    const COLOR_RED = 0x00;

    const COLOR_BLACK = 0x01;

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
    const OleDirectoryEntrySize = 0x80;

    // Unpack() format string for the 436 byte DIFAT array in the header
    const OleHeaderDIFATFormat = 'V109';                

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

    public function new($filepath, $format = true, $version = 0x0003, $clsid = null)
    {
        $this->filepath = $filepath;
        if ($format) {
            $this->format($version, $clsid);
        }
        return $this;
    }

    public function open($source)
    {
        if (is_resource($source) && get_resource_type($source) === 'stream') {
            $this->CreateFromStream($source);
        } elseif (is_string($source)) {
            if (strlen($source) >= 8 && unpack('H16', $source)[1] === OleHeader::MAGIC) {
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

    private function serializeHeaderDIFAT()
    {
        return pack(self::OleHeaderDIFATFormat, ...array_slice($this->DIFAT, 0, 109));
    }

    private function writeHeader()
    {
        if (!$this->header->isDirty()) {
            return;
        }

        $headerdata = $this->header->serialize() . $this->serializeHeaderDIFAT();

        rewind($this->stream);
        fwrite($this->stream, $headerdata);
        if ($this->sectorsize > self::OleHeaderSize) {
            fwrite($this->stream, str_repeat(chr(0), $this->sectorsize - self::OleHeaderSize)); // pad header to a full sector
        }

        $this->header->makeClean();
    }

    private function writeDIFAT()
    {
        for ($difat = array_slice($this->DIFAT, 109), $i = 1; $difatdata = array_splice($difat, 0, $this->sectorsize/4 - 1); $i++) {
            if ($i > count($this->DIFATSectors)) {
                throw new \BadMethodCallException('number of DIFAT entries doesn\'t match number of DIFAT sectors');
            }

            if ($i === count($this->DIFATSectors)) {
                $difatdata[] = self::ENDOFCHAIN;
            } else {
                $difatdata[] = $this->DIFATSectors[$i];
            }
            $data = pack($this->FAT_sectorformat, $difatdata);
            $this->setSectorData($this->DIFATSectors[$i-1], $data);
        }
    }

    private function writeFAT()
    {
        if (!$this->FATDirty) {
            return;
        }
        
        for ($fat = $this->FAT, $i = 0; $fatdata = array_splice($fat, 0, $this->sectorsize / 4); $i++) {
            $fatstring = pack($this->FAT_sectorformat, ...$fatdata);
            if ($this->DIFAT[$i] >= self::MAXREGSECT) {
                throw new \BadMethodCallException('Number of FAT entries does not correspond to number of FAT sectors in DIFAT');
            }
            $this->setSectorData($this->DIFAT[$i], $fatstring);
        }

        $this->FATDirty = false;
    }

    private function writeDirectorySectors()
    {
        $s = $this->header->getFirstDirectorySector();
        if ($s === self::ENDOFCHAIN) {
            return;
        }

        $data = '';
        foreach ($this->directoryEntries as $index => $directoryEntry) {
            $data .= $directoryEntry->serialize();
            if (strlen($data) === $this->sectorsize) {
                $this->setSectorData($s, $data);
                $s = $this->FAT[$s];
                if ($s === self::ENDOFCHAIN && $index < count($this->directoryEntries) - 1) {
                    throw new \BadMethodCallException('Insufficient directory sectors allocated to hold all directory entries');
                }
                $data = '';
            }
        }

        if (strlen($data) < $this->sectorsize) {
            $data = str_pad($data, $this->sectorsize, OleDirectoryEntry::EMPTY);
            $this->setSectorData($s, $data);
        }
    }

    public function writeMiniFAT()
    {
        if (!$this->miniFAT) {
            return;
        }

        for ($i = 0, $s = $this->header->getFirstMiniFATSector(), $fat = $this->miniFAT; $fatdata = array_splice($fat, 0, $this->sectorsize / 4); $i++, $s = $this->FAT[$s]) {
            $fatstring = pack($this->FAT_sectorformat, ...$fatdata);
            if ($s >= self::MAXREGSECT) {
                throw new \BadMethodCallException('Number of miniFAT entries does not correspond to number of miniFAT sectors in the FAT');
            }
            $this->setSectorData($s, $fatstring);
        }
    }

    private function writeDirtySectors()
    {
        foreach ($this->dirtySectors as $sector => $data) {
            $this->writeSectorData($sector, $data);
        }
        $this->dirtySectors = [];
    }

    public function save($dest = null)
    {
        $oldstream = null;
        if ($dest === null) {
            $dest = $this->filepath;
        } elseif (!$this->filepath) {
            $this->filepath = $dest;
        } elseif ($dest !== $this->filepath) {
            $oldstream = $this->stream;
            $this->stream = null;
        }

        if (!$this->stream) {
            $this->stream = fopen($dest, 'w+');
        }

        $this->writeHeader();
        $this->writeDIFAT();
        $this->writeFAT();
        $this->writeMiniFAT();
        $this->writeDirectorySectors();
        $this->writeDirtySectors();

        if ($oldstream) {
            fclose($this->stream);
            $this->stream = $oldstream;
        }
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
        $this->directoryEntries = null;
        return $this;
    }

    private function initialize()
    {
        $this->sectorsize = 1 << $this->header->getSectorShift();
        $this->minisectorsize = 1 << $this->header->getMiniSectorShift();
        $this->miniStreamCutoff = $this->header->getMiniStreamCutoff();
        switch ($this->header->getVersion()) {
            case 0x03:
                $this->FAT_sectorformat = self::OleV3FATSectorFormat;
                $this->DIFAT_sectorformat = self::OleV3DIFATSectorFormat;
                break;
            case 0x04:
                $this->FAT_sectorformat = self::OleV4FATSectorFormat;
                $this->DIFAT_sectorformat = self::OleV4DIFATSectorFormat;
                if ($this->stream) {
                    fseek($this->stream, $this->sectorsize, SEEK_SET); // move to start of first sector after header
                }
                break;
            default:
                throw new \Exception("Invalid OLE version");
        }
    }

    private function initializeHeader($version)
    {
        $this->header = OleHeader::new($this, $version);
        $this->initialize();
    }

    private function intitializeFAT()
    {
        // create one sector worth of fat entries, and reserve sector 0 for the new FAT sector
        $this->DIFAT = array_fill(0, 109, self::FREESECT);
        $this->FAT = [];
        $this->allocateFATSector();
    }

    private function initializeRootStorage($clsid = null)
    {
        // TODO: check that first directory sector isn't already allocated
        $this->allocateDirectorySector();
        $this->rootEntry = OleDirectoryEntry::new($this, 0, 'Root Entry', self::RootStorageObject, $clsid);
        $this->directoryEntries = [$this->rootEntry];
        $this->rootStorage = new OleStorage($this);
    }

    public function initializeMiniStream()
    {
        // TODO: make sure the ministream isn't already initialized
        $this->header->setFirstMiniFATSector($this->allocateSector());
        $this->miniFAT += array_fill(0, $this->sectorsize / 4, self::FREESECT);
        $this->rootEntry->setStartingSector($this->allocateSector());
        $this->rootEntry->setStreamSize($this->sectorsize);
    }

    public function format($version = 0x0003, $clsid = null)
    {
        if ($this->stream) {
            ftruncate($this->stream, 0);
        }
        $this->header = null;
        $this->FAT = [];
        $this->miniFAT = [];
        $this->rootStorage = null;
        $this->directoryEntries = null;

        // OLE Spec says the minimum size OLE file has a header, one FAT sector, and one directory sector
        $this->initializeHeader($version);
        $this->intitializeFAT();
        $this->initializeRootStorage($clsid);
    }

    private function allocateDirectorySector()
    {
        if (($s = $this->header->getFirstDirectorySector()) === self::ENDOFCHAIN) {
            return $this->header->setFirstDirectorySector($this->allocateSector());
        } else {
            while ($this->FAT[$s] !== self::ENDOFCHAIN) {
                $s = $this->FAT[$s];
            }
            return $this->allocateSector(self::ENDOFCHAIN, $s);
        }
    }

    private function allocateDIFATSector()
    {
        $s = $this->allocateSector(self::DIFSECT);
        $this->DIFAT = array_merge($this->DIFAT, array_fill(0, $this->sectorsize / 4 - 1, self::FREESECT));
        if (!$this->DIFATSectors) {
            $this->DIFATSectors = [];
            $this->header->setFirstDIFATSector($s);
        }
        $this->header->incDIFATSectorCount();
        $this->DIFATSectors[] = $s;
    }

    private function allocateFATSector()
    {
        $result = 1;
        $i = count($this->FAT);
        $this->FAT = array_merge($this->FAT, array_fill(0, $this->sectorsize / 4, self::FREESECT));
        $this->FAT[$i] = self::FATSECT;
        $this->header->incFATSectorCount();

        // Add a DIFAT entry for the new FAT sector, allocating a new DIFAT sector if necessary
        for ($d = 0; $d < count($this->DIFAT) && $this->DIFAT[$d] !== self::FREESECT; $d++);
        if ($d === count($this->DIFAT)) {
            $this->allocateDIFATSector();
            $result++;
        }
        $this->DIFAT[$d] = $i;

        $this->FATDirty = true;

        // return the total number of new sectors allocated (1 for the new FAT sector and possibly 1 for a new DIFAT sector)
        return $result;
    }

    private function allocateSector($code = self::ENDOFCHAIN, $prevSector = null)
    {
        for ($i = 0; $i < count($this->FAT) && $this->FAT[$i] !== self::FREESECT; $i++);
        if ($i ===  count($this->FAT)) {
            $i += $this->allocateFATSector();
        }
        
        $this->FAT[$i] = $code;
        if ($prevSector) {
            $this->FAT[$prevSector] = $i;
        }
        $this->dirtySectors[$i] = str_repeat(chr(0), $this->sectorsize);
        $this->FATDirty = true;
        return $i;
    }

    private function allocateMiniSector($code = self::ENDOFCHAIN, $prevSector = null)
    {
        if (!$this->miniFAT) {
            $this->initializeMiniStream();
        }

        for ($i = 0; $i < count($this->miniFAT) && $this->miniFAT[$i] !== self::FREESECT; $i++);
        if ($i ===  count($this->miniFAT)) {
            // Alloate a new mini-FAT sector
            $this->allocateSector(self::ENDOFCHAIN, $this->endOfChain($this->header->getFirstMiniFATSector()));
            $this->miniFAT = array_merge($this->miniFAT, array_fill(0, $this->sectorsize / 4, self::FREESECT));
        }
        $this->miniFAT[$i] = $code;
        while ($i * $this->minisectorsize >= $this->getMiniStreamSize()) {
            // allocate a new sector to the ministream
            $this->allocateSector(self::ENDOFCHAIN, $this->endOfChain($this->rootEntry->getStartingSector()));
            $this->rootEntry->setStreamSize($this->rootEntry->getStreamSize() + $this->sectorsize);
        }
        if ($prevSector) {
            $this->miniFAT[$prevSector] = $i;
        }
        
        return $i;
    }

    /**
     * Whether stream's size is below the threshold for being stored in the ministream storage
     * 
     * @param int|OleDirectoryEntry $stream
     */
    public function isMiniStream($stream)
    {
        if (!$stream instanceof OleDirectoryEntry) {
            $stream = $this->directoryEntries[$stream];
        }
        return $stream->getObjectType() === self::StreamObject 
            && $stream->getStreamSize() < $this->header->getMiniStreamCutoff();
    }

    /**
     * Get the size of big sectors in this Ole file (512 or 4096)
     */
    public function getSectorSize()
    {
        return $this->sectorsize;
    }

    /**
     * Get the size of big sectors in this Ole file (512 or 4096)
     */
    public function getMiniSectorSize()
    {
        return $this->minisectorsize;
    }

    /**
     * Get the the sector number of the next sector in the chain
     *
     * @param int $sector
     */
    public function getNextSector($sector)
    {
        if ($sector <= self::MAXREGSECT && $sector < count($this->FAT)) {
            return $this->FAT[$sector];
        } else {
            return $sector;
        }
    }

    /**
     * Get the the sector number of the next sector in a chain in the MiniFAT
     *
     * @param int $sector
     */
    public function getNextMiniSector($sector)
    {
        if ($sector <= self::MAXREGSECT && $sector < count($this->miniFAT)) {
            return $this->miniFAT[$sector];
        } else {
            return $sector;
        }
    }

    private function endOfChain($sector, $fat = null)
    {
        if ($fat === null) {
            $fat = $this->FAT;
        }

        if ($sector <= self::MAXREGSECT && $sector < count($fat)) {
            for ($s = $sector; $fat[$s] !== self::ENDOFCHAIN; $s = $fat[$s]);
            return $s;
        } else {
            return false;
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
        if (array_key_exists($sector, $this->dirtySectors)) {
            $data = $this->dirtySectors[$sector];
        } else {
            fseek($this->stream, ($sector + 1) * $this->sectorsize);
            $data = fread($this->stream, $this->sectorsize);

            if (!$data)
                throw new \Exception("Could not read sector {$sector}");
        }
        return $data;
    }

    /**
     * Read a single sector from the underlying file descriptor
     *
     * @param int $sector
     * @throws \Exception
     * @return string
     */
    public function setSectorData($sector, $data)
    {
        if (strlen($data) < $this->sectorsize) {
            $data = str_pad($data, $this->sectorsize, chr(0));
        }

        $this->dirtySectors[$sector] = $data;
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
     * Get the currently allocated size of the ministream
     * 
     * @return int
     */
    private function getMiniStreamSize()
    {
        return $this->rootEntry->getStreamSize();
    }
    /**
     * Read a single sector from the mini stream
     *
     * @param int $sector
     * @return string
     */
    public function getMiniSectorData($sector)
    {
        return $this->read(0, $this->minisectorsize, $sector * $this->minisectorsize);
    }

    /**
     * Read a single sector from the underlying file descriptor
     *
     * @param int $sector
     * @throws \Exception
     * @return string
     */
    public function setMiniSectorData($sector, $data)
    {
        if (strlen($data) < $this->minisectorsize) {
            $data = str_pad($data, $this->minisectorsize, chr(0));
        }

        $this->write(0, $data, $sector * $this->minisectorsize);
        return $data;
    }

    /**
     * Read the header record and initialize internal data structures/values
     */
    private function readHeader()
    {
        $data = (string) fread($this->stream, self::OleHeaderSize);
        
        if (!$data) {
            throw new \Exception("Could not read header sector");
        }

        $this->header = new OleHeader($this, substr($data, 0, 76));
        $this->initialize();

        $this->DIFAT = array_values(unpack(self::OleHeaderDIFATFormat, substr($data, 76)));
    }

    /** 
     * Load the DIFAT sectors beyond the DIFAT entries stored in the header
     * and add them to the DIFAT array
     */
    private function readDIFAT()
    {
        $s = $this->header->getFirstDIFATSector();
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
        $s = $this->header->getFirstMiniFATSector();
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
        for ($i = 0; $i < $this->sectorsize; $i += self::OleDirectoryEntrySize) {
            // if the first Unicode character of the entry name is null, then the entry is unused
            // meaning we've reached the end of the list of directory entries
            if ($data[$i] === chr(0) && $data[$i+1]=== chr(0)) {
                break;
            }
            
            $objectId = count($this->directoryEntries);
            $this->directoryEntries[] = new OleDirectoryEntry($this, $objectId, $data, $i);
        }
    }

    /**
     * Read the all directoryEntry entries for the Ole file
     */
    private function readDirectoryEntries()
    {
        $this->directoryEntries = [];
        $s = $this->header->getFirstDirectorySector();
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
        $this->rootStorage = new OleStorage($this);
    }

    /**
     * Read $bytes bytes from the stream corresponding to streamid starting at $offset
     *
     * @param int $streamid
     * @param int $bytes
     * @param int $offset
     * @throws \InvalidArgumentException
     * @return string
     */
    public function read($streamid, $bytes = -1, $offset = 0)
    {
        if ($streamid < 0 || $streamid >= sizeof($this->directoryEntries))
            throw new \InvalidArgumentException("Invalid StreamID $streamid");

        if ($this->directoryEntries[$streamid]->getObjectType() !== self::StreamObject 
            && $this->directoryEntries[$streamid]->getObjectType() !== self::RootStorageObject)
            throw new \InvalidArgumentException("StreamID $streamid is not a stream");

        if ($offset > $this->directoryEntries[$streamid]->getStreamSize())
            throw new \InvalidArgumentException("Attempt to read past end of stream");

        // $bytes = -1 means read the whole stream
        if ($bytes == -1) {
            $bytes = $this->directoryEntries[$streamid]->getStreamSize();
        }

        if ($streamid === 0 || $this->directoryEntries[$streamid]->getStreamSize() >= $this->miniStreamCutoff) {
            $readsector = array(
                $this,
                'getSectorData'
            );
            $fat =& $this->FAT;
            $bs = $this->sectorsize;
        } else {
            $readsector = array(
                $this,
                'getMiniSectorData'
            );
            $fat =& $this->miniFAT;
            $bs = 64;
        }

        // first find the sector containing the starting offset
        $s = $this->directoryEntries[$streamid]->getStartingSector();
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
     * Read $bytes bytes from the stream corresponding to streamid starting at $offset
     *
     * @param int $streamid
     * @param string $data
     * @param int $offset
     * @throws \InvalidArgumentException
     * @return int - bytes written
     */
    public function write($streamid, $data, $offset = 0)
    {
        if ($streamid < 0 || $streamid >= sizeof($this->directoryEntries))
            throw new \InvalidArgumentException("Invalid StreamID $streamid");

        $entry = $this->directoryEntries[$streamid];
        if ($entry->getObjectType() !== self::StreamObject && $entry->getObjectType() !== self::RootStorageObject)
            throw new \InvalidArgumentException("StreamID $streamid is not a stream");

        // TODO: move a ministream allocated stream to the regular FAT if it gets too big for the ministream
        if ($streamid == 0 || max($entry->getStreamSize(), $offset + strlen($data)) >= $this->miniStreamCutoff) {
            $readsector = array(
                $this,
                'getSectorData'
            );
            $writesector = array(
                $this,
                'setSectorData'
            );
            $allocatesector = array(
                $this,
                'allocateSector'
            );
            $fat =& $this->FAT;
            $bs = $this->sectorsize;
        } else {
            $readsector = array(
                $this,
                'getMiniSectorData'
            );
            $writesector = array(
                $this,
                'setMiniSectorData'
            );
            $allocatesector = array(
                $this,
                'allocateMiniSector'
            );
            $fat =& $this->miniFAT;
            $bs = 64;
        }

        // first find the sector containing the starting offset
        // if offset is beyond the current end of the stream, allocate additional empty sectors
        if (($s = $this->directoryEntries[$streamid]->getStartingSector()) === self::ENDOFCHAIN) {
            $s = $allocatesector();
            $this->directoryEntries[$streamid]->setStartingSector($s);
        }
        $i = 0;
        while ($i < $offset) {
            if ($offset < $i + $bs)
                break;
            $i += $bs;
            if ($fat[$s] === self::ENDOFCHAIN) {
                $s = $allocatesector(self::ENDOFCHAIN, $s);
            } else {
                $s = $fat[$s];
            }
        }
        
        // replace from the starting offset to the end of the starting sector
        $strOffset = 0;
        $bytesLeft = $byteswritten = strlen($data);
        $l = min($bytesLeft, $bs - ($offset % $bs));
        $sectorData = substr_replace($readsector($s), substr($data, $strOffset, $l), $offset - $i, $l);
        $writesector($s, $sectorData);
        $strOffset += $l;
        $bytesLeft -= $l;

        // keep writing sectors (or a partial sector at the end) until we've written all the data
        while ($bytesLeft) {
            if ($fat[$s] === self::ENDOFCHAIN) {
                $s = $allocatesector(self::ENDOFCHAIN, $s);
            } else {
                $s = $fat[$s];
            }
            $l = min($bytesLeft, $bs);
            $sectorData = substr_replace($readsector($s), substr($data, $strOffset, $l), 0, $l);
            $writesector($s, $sectorData);
            $strOffset += $l;
            $bytesLeft -= $l;
        }
        
        if ($offset + $byteswritten > $entry->getStreamSize()) {
            $entry->setStreamSize($offset + $byteswritten);
        }

        return $byteswritten;
    }

    /**
     * Read all of the data for a stream within the Ole file
     *
     * @param int $streamid
     * @return string
     */
    public function getStreamData($streamid)
    {
        if ($streamid instanceof OleDirectoryEntry) {
            $entry = $streamid;
        } else {
            $entry = $this->directoryEntries[$streamid];
        }
        if ($entry->getObjectType() === 1) {
            return null; // should this throw an error?
        }

        if ($entry->getId() === 0 || $entry->getStreamSize() >= $this->miniStreamCutoff) {
            $readsector = array(
                $this,
                'getSectorData'
            );
            $fat = $this->FAT;
        } else {
            $readsector = array(
                $this,
                'getMiniSectorData'
            );
            $fat = $this->miniFAT;
        }

        $s = $entry->getStartingSector();
        $data = '';
        while ($s != self::ENDOFCHAIN) {
            $data .= $readsector($s);
            $s = $fat[$s];
        }

        return substr($data, 0, $entry->getStreamSize());
    }

    /**
     * Return the streamid for the main document stream in this Ole file or false if none could be found
     *
     * @return number|boolean
     */
    public function getDocumentStream()
    {
        foreach ($this->directoryEntries as $id => $stream) {
            if ($stream->getObjectType() === 2 && $stream->getStartingSector() === 0)
                return $id;
        }

        return false;
    }

    /**
     * Add an OleDirectoryEntry to the document's directory entry list, and optionally insert it into a storage
     * 
     * @param OleDirectoryEntry $entry
     * @param OleStorage|bool|null $parent - false means don't insert, null means default to the root storage
     * @param int $objectId - the objectId to assign or null to assign one here
     * @return int - the assigned objectId
     */
    public function addEntry(OleDirectoryEntry $entry, $parent = null, $objectId = null)
    {
        $entriesPerSector = $this->sectorsize / self::OleDirectoryEntrySize;
        if (count($this->directoryEntries) % $entriesPerSector === 0) {
            $this->allocateDirectorySector();
        }

        if ($parent === null) {
            $parent = $this->rootStorage;
        }
        if (!$objectId) {
            $objectId = count($this->directoryEntries);
        }
        $this->directoryEntries[$objectId] = $entry;
        if ($parent) {
            $parent->insertEntry($entry);
        }
        return $objectId;
    }

    public function newEntry($name, $type = self::StreamObject, $clsid = null, OleStorage $parent = null) 
    {
        $objectId = count($this->directoryEntries);
        $entry = OleDirectoryEntry::new($this, $objectId, $name, $type, $clsid);
        $this->addEntry($entry, $parent, $objectId);

        return $entry;
    }

    public function getObject($entry)
    {
        if (is_null($entry)) {
            $entryId = $this->getDocumentStream();
            $entry = $this->directoryEntries[$entryId];
        } elseif (is_string($entry)) {
            if (!$entryId = $this->findEntryByName($entry)) {
                throw new \Exception("Stream {$entry} not found");
            }
            $entry = $this->directoryEntries[$entry];
        } elseif (is_int($entry)) {
            $entry = $this->directoryEntries[$entry];
        } elseif (!$entry instanceof OleDirectoryEntry) {
            throw new \Exception("Invalid stream {$entry}");
        }

        if ($entry->getObjectType() === self::RootStorageObject) {
            return $this->rootStorage;
        } else {
            $class = self::TYPE_MAP[$entry->getObjectType()];
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
        foreach ($this->directoryEntries as $id => $entry) {
            if ($entry->getName() === $entryName) {
                return $id;
            }
        }

        return false;
    }

    /**
     * Return the root directory of the Ole document
     *
     * @return OleStorage
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
        $this->readDirectoryEntries();
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
        return new \ArrayIterator($this->directoryEntries);
    }

    public function count()
    {
        return count($this->directoryEntries);
    }

    public function offsetGet($offset)
    {
        return $this->directoryEntries[$offset];
    }

    public function offsetExists($offset)
    {
        return array_key_exists($offset, $this->directoryEntries);
    }

    public function offsetUnset($offset)
    {
        throw new \BadMethodCallException('OleDocument file specs are readonly');
    }

    public function offsetSet($offset, $value)
    {
        throw new \BadMethodCallException('OleDocument file specs are readonly');
    }

    public function getMaxUsedSector()
    {
        $result = 0;
        foreach ($this->FAT as $index => $fatentry)
        {
            if ($fatentry !== self::FREESECT) {
                $result = $index;
            }
        }
        return $result;
    }

    public function serialize()
    {
        $this->writeDIFAT();
        $this->writeFAT();
        $this->writeMiniFAT();
        $this->writeDirectorySectors();

        $result = $this->header->serialize() . $this->serializeHeaderDIFAT();
        if ($this->sectorsize > self::OleHeaderSize) {
            $result .= str_repeat(chr(0), $this->sectorsize - self::OleHeaderSize);
        }
        for ($i = 0; $i <= $this->getMaxUsedSector(); $i++) {
            $result .= $this->getSectorData($i);
        }
        return $result;
    }
}
