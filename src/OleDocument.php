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
     * Block size for this file: 512 for version <= 3, 4096 for version 4
     *
     * @var int
     */
    private $blocksize;

    /**
     * Format string of regular blocks for unpack()
     * Either OleV3FATBlockFormat or OleV4FATBlockFormat
     *
     * @var string
     */
    private $FAT_blockformat;

    /**
     * Format string of directory information blocks for unpack()
     * Either OleV3DIFATBlockFormat or OleV4DIFATBlockFormat
     *
     * @var string
     */
    private $DIFAT_blockformat;

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
    private $MiniFAT;

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

    // Unpack() format string for a 512 byte FAT block
    // Each FAT block contains 128 FAT entries
    const OleV3FATBlockFormat = 'V128';

    // Unpack() format string for a 512 byte directory information FAT block
    // Each block contains 127 FAT entries for blocks containing directory information and
    // one entry pointing to the next directory information FAT block
    const OleV3DIFATBlockFormat = 'V127FATBlocks/V1NextDIFATBlock';

    // Unpack() format string for a 4096 byte FAT block
    // Each FAT block contains 1024 FAT entries
    const OleV4FATBlockFormat = 'V1024';

    // Unpack() format string for a 4096 byte directory information FAT block
    // Each block contains 1023 FAT entries for blocks containing directory information and
    // one entry pointing to the next directory information FAT block
    const OleV4DIFATBlockFormat = 'V1023FATBlocks/V1NextDIFATBlock';

    // @formatter:off
    // Unpack() format string for the Ole Header
    const OleHeaderFormat =
            'H16Signature/' .           # 00 8 bytes
            'H32CLSID/' .               # 08 16 bytes
            'v1MinorVersion/' .         # 18 2 bytes
            'v1MajorVersion/' .         # 1A 2 bytes
            'v1ByteOrder/' .            # 1C 2 bytes
            'v1BlockShift/' .           # 1E 2 bytes
            'v1MiniBlockShift/' .       # 20 2 bytes
            'Z6Reserved1/' .            # 22 6 bytes
            'V1DirectoryBlocks/' .      # 28 4 bytes
            'V1FATBlocks/' .            # 2C 4 bytes
            'V1FirstDirectoryBlock/' .  # 30 4 bytes
            'V1TransactionSignature/' . # 34 4 bytes
            'V1MiniStreamCutoff/' .     # 38 4 bytes
            'V1FirstMiniFATBlock/' .    # 3C 4 bytes
            'V1MiniFATBlockCount/' .    # 40 4 bytes
            'V1FirstDIFATBlock/' .      # 44 4 bytes
            'V1DIFATBlockCount/' .      # 48 4 bytes
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
            'V1StartingBlock/' .    # 74 4 bytes
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
        $this->Close();
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

    }

    /**
     * Close the underlying file resource and reset the internal structures
     */
    public function close()
    {
        if ($this->stream)
            fclose($this->stream);
        $this->stream = null;
        $this->FAT = null;
        $this->MiniFAT = null;
        $this->rootStorage = null;
        $this->fileSpecs = null;
    }

    /**
     * Get the size of big blocks in this Ole file (512 or 4096)
     */
    public function getBlocksize()
    {
        return $this->blocksize;
    }

    /**
     * Get the the block number of the next block in the chain
     *
     * @param int $block
     */
    public function getNextBlock($block)
    {
        if (($block & 0xFFFFFFF8) !== 0xFFFFFFF8) {
            return $this->FAT[$block];
        } else {
            return $block;
        }
    }

    /**
     * Read a single block from the underlying file descriptor
     *
     * @param int $block
     * @throws \Exception
     * @return string
     */
    public function getBlockData($block)
    {
        fseek($this->stream, ($block + 1) * $this->blocksize);
        $data = fread($this->stream, $this->blocksize);

        if (!$data)
            throw new \Exception("Could not read block {$block}");

        return $data;
    }

    /**
     * Read a seingle block from the mini stream
     *
     * @param int $block
     * @return string
     */
    public function getMiniBlockData($block)
    {
        return $this->read(0, 64, $block * 64);
    }

    /**
     * Load the entire FAT into $this->FAT
     */
    private function readFAT()
    {
        $this->FAT = array();
        for ($i = 0; $i < 109; $i++) {
            if ($this->header['DIFAT'][$i] >= self::ENDOFCHAIN)
                return;
            $data = $this->getBlockData($this->header['DIFAT'][$i]);
            $entries = unpack($this->FAT_blockformat, $data);
            $this->FAT = array_merge($this->FAT, $entries);
        }

        $s = $this->header['FirstDIFATBlock'];
        while ($s != self::ENDOFCHAIN) {
            $data = $this->getBlockData($s);
            $difat = unpack($this->DIFAT_blockformat, $data);
            $difat['FATBlocks'] = array_values(array_splice($difat, 0, sizeof($difat) - 1));

            for ($i = 0; $i < $this->blocksize / 4 - 1; $i++) {
                if ($difat['FATBlocks'][$i] == self::ENDOFCHAIN)
                    return;
                $data = $this->getBlockData($this->header['DIFAT'][$i]);
                $entries = unpack($this->FAT_blockformat, $data);
                $this->FAT = array_merge($this->FAT, $entries);
            }

            $s = $difat['NextDIFATBlock'];
        }
    }

    /**
     * Load the entire MiniFAT into $this->MiniFAT
     */
    private function readMiniFAT()
    {
        $this->MiniFAT = array();
        $s = $this->header['FirstMiniFATBlock'];
        while ($s != self::ENDOFCHAIN) {
            $data = $this->getBlockData($s);
            $entries = unpack($this->FAT_blockformat, $data);
            $this->MiniFAT = array_merge($this->MiniFAT, $entries);
            $s = $this->FAT[$s];
        }
    }

    /**
     * Read directory entries from a directory information block and add them to an existing
     * array, if passed.
     * Return an array containing the directory entries.
     *
     * @param int $block
     * @param array $entries
     * @return string|string[]
     */
    private function readDirectoryBlock($block)
    {
        $data = $this->getBlockData($block);
        for ($i = 0; $i < $this->blocksize / 128; $i++) {
            $newentry = unpack(OleDocument::OleDirectoryEntryFormat, $data, $i * OleDocument::OleDirectoryEntrySize);
            // unpack cuts off the final byte of the final UTF-16LE character if it is null, so we have to add it back on
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
        $s = $this->header['FirstDirectoryBlock'];
        while ($s != OleDocument::ENDOFCHAIN) {
            $this->readDirectoryBlock($s);
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
            $readblock = array(
                $this,
                'getBlockData'
            );
            $fat = $this->FAT;
            $bs = $this->blocksize;
        } else {
            $readblock = array(
                $this,
                'getMiniBlockData'
            );
            $fat = $this->MiniFAT;
            $bs = 64;
        }

        // first find the block containing the starting offset
        $s = $this->fileSpecs[$streamid]['StartingBlock'];
        $i = 0;
        while ($s != self::ENDOFCHAIN && $i < $offset) {
            if ($offset < $i + $bs)
                break;
            $i += $bs;
            $s = $fat[$s];
        }

        // grab from the starting offset to the end of the starting block
        $data = substr($readblock($s), $offset - $i, $bs - ($offset % $bs));

        // keep adding blocks until we've got all the bytes requested
        if ($s != self::ENDOFCHAIN) {
            $s = $fat[$s];
            while (strlen($data) < $bytes && $s != self::ENDOFCHAIN) {
                $data .= $readblock($s);
                $s = $fat[$s];
            }
        }

        // If the last block read added more bytes than requested, truncate the returned data
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
    public function getData($streamid)
    {
        if ($this->fileSpecs[$streamid]['ObjectType'] != 2)
            return null; // should this throw an error?

        if ($streamid == 0 || $this->fileSpecs[$streamid]['StreamSize'] >= $this->header['MiniStreamCutoff']) {
            $readblock = array(
                $this,
                'getBlockData'
            );
            $fat = $this->FAT;
            // $bs = $this->blocksize;
        } else {
            $readblock = array(
                $this,
                'getMiniBlockData'
            );
            $fat = $this->MiniFAT;
            // $bs = 64;
        }

        $s = $this->fileSpecs[$streamid]['StartingBlock'];
        $data = '';
        while ($s != self::ENDOFCHAIN) {
            $data .= $readblock($s);
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
            if ($stream['ObjectType'] == 2 && $stream['StartingBlock'] == 0)
                return $id;
        }

        return false;
    }

    public function getStream($stream)
    {
        if (is_null($stream)) {
            $stream = $this->getDocumentStream();
            $filespec = $this->fileSpecs[$stream];
        } elseif (is_string($stream)) {
            if (!$stream = $this->FindStreamByName($stream)) {
                throw new \Exception("Stream {$stream} not found");
            }
            $filespec = $this->fileSpecs[$stream];
        } elseif (is_int($stream)) {
            $filespec = $this->fileSpecs[$stream];
        } elseif (is_array($stream)) {
            $filespec = $stream;
        } else {
            throw new \Exception("Invalid stream {$stream}");
        }

        if ($filespec['ObjectType'] === self::RootStorageObject) {
            return $this->rootStorage;
        } else {
            $class = self::TYPE_MAP[$filespec['ObjectType']];
            return new $class($this, $stream);
        }
    }

    /**
     * Return the streamid for the stream with the passed stream name
     *
     * @param string $streamName
     * @return number|boolean
     */
    public function findStreamByName($streamName)
    {
        // could use the red/black tree to do this more quickly
        foreach ($this->fileSpecs as $id => $stream) {
            if ($stream['EntryName'] == $streamName) {
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

        $data = (string) fread($strm, self::OleHeaderSize);
        if (!$data) {
            throw new \Exception("Could not read header block");
        }

        $this->header = unpack($this::OleHeaderFormat, $data);
        if ($this->header['Signature'] != self::OleSignature) {
            throw new \Exception("Stream is not an Ole file");
        }

        $this->header['DIFAT'] = array_values(array_splice($this->header, -109, 109));
        switch ($this->header['MajorVersion']) {
            case 0x03:
                $this->blocksize = 512;
                $this->FAT_blockformat = self::OleV3FATBlockFormat;
                $this->DIFAT_blockformat = self::OleV3DIFATBlockFormat;
                break;
            case 0x04:
                $this->blocksize = 4096;
                $this->FAT_blockformat = self::OleV4FATBlockFormat;
                $this->DIFAT_blockformat = self::OleV4DIFATBlockFormat;
                fseek($this->stream, 4096, SEEK_SET); // move to start of first block after header
                break;
            default:
                throw new \Exception("Invalid BlockShift");
        }

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

