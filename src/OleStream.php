<?php
namespace Cryptodira\PhpOle;

/**
 * File-like structure for rading from an individual stream within an Ole compound document
 *
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
class OleStream extends OleObject
{

    private $fat;

    private $sectorsize;

    private $readsector;

    private $fatpointer;

    private $firstsector;

    private $buffer;

    private $size;

    private $pos;

    /**
     * Create a new Ole stream object for a given stream within an Ole compound document
     * This class buffers reads by loading one full sector at a time to make small reads more efficient
     *
     * @param OleDocument $root
     * @param mixed $stream
     * @throws \Exception
     */
    public function __construct($root, $stream = null)
    {
        parent::__construct($root, $stream);

        $this->firstsector = $this->entry->getStartingSector();
        $this->size = $this->entry->getStreamSize();

        // Use a closure/binding to access the internals of the Ole document -- simulating a C++ friend relationship
        $initialize = function (OleDocument $root, $size, &$readsector, &$fat, &$sectorsize) {
            if ($size >= $root->header['MiniStreamCutoff']) {
                $readsector = \Closure::fromCallable([
                    $root,
                    'getSectorData'
                ]);
                $fat = $root->FAT;
                $sectorsize = $root->sectorsize;
            } else {
                $readsector = \Closure::fromCallable([
                    $root,
                    'getMiniSectorData'
                ]);
                $fat = $root->MiniFAT;
                $sectorsize = 64;
            }
        };

        $initialize = $initialize->bindTo($this, $root);
        $initialize($root, $this->size, $this->readsector, $this->fat, $this->sectorsize);

        $this->fatpointer = $this->firstsector;
        $this->pos = 0;
        $this->buffer = null;
    }

    /**
     * Move the file pointer to a specific position within the stream
     *
     * @param int $pos
     */
    private function doSeek($pos)
    {
        $oldsector = intdiv($this->pos, $this->sectorsize);
        $newsector = intdiv($pos, $this->sectorsize);

        if ($newsector != $oldsector) {
            $this->fatpointer = $this->firstsector;
            for ($i = 0; $i < $newsector && $this->fatpointer != OleDocument::ENDOFCHAIN; $i++)
                $this->fatpointer = $this->fat[$this->fatpointer];

            $this->buffer = null;
        }

        $this->pos = $pos;
    }

    /**
     * Check whether file pointer is at the end of the stream data
     *
     * @return boolean
     */
    public function eof()
    {
        return ($this->pos == $this->size);
    }

    /**
     * Reset the file pointer to the beginning of the stream
     */
    public function rewind()
    {
        $this->buffer = null;
        $this->fatpointer = $this->firstsector;
        $this->pos = 0;
    }

    /**
     * Move the file pointer to the passed position
     *
     * @param int $pos
     * @param int $seektype
     * @return int - 0 on success or -1 if the seek would move the file pointer beyond the end of the stream
     */
    public function seek($pos, $seektype = SEEK_SET)
    {
        switch ($seektype) {
            case SEEK_SET:
                break;
            case SEEK_CUR:
                $pos = $this->pos + $pos;
                break;
            case SEEK_END:
                $pos = $this->size + $pos;
                break;
        }

        if ($pos > $this->size) {
            $this->pos = $this->size;
            $this->buffer = null;
            return -1;
        }

        if ($pos != $this->pos)
            $this->doSeek($pos);

        return 0;
    }

    /**
     * Return the current file pointer position within the stream
     *
     * @return int
     */
    public function tell()
    {
        return $this->pos;
    }

    /**
     * Read a maximum of $bytes bytes from the stream at the current position
     *
     * @param int $bytes
     * @return string
     */
    public function read($bytes = null)
    {
        if ($this->pos == $this->size) {
            return '';
        }

        if ($bytes === null) {
            $bytes = $this->size - $this->pos;
        }

        if (is_null($this->buffer)) {
            $this->buffer = ($this->readsector)($this->fatpointer);
        }

        $sector_offset = $this->pos % $this->sectorsize;
        if ($bytes <= $this->sectorsize - $sector_offset) {
            $data = substr($this->buffer, $sector_offset, $bytes);
            $this->pos += $bytes;
            if ($sector_offset + $bytes == $this->sectorsize) {
                $this->fatpointer = $this->fat[$this->fatpointer];
                $this->buffer = null;
            }
        } else {
            $data = substr($this->buffer, $sector_offset);
            $this->buffer = null;
            $bytesread = $this->sectorsize - $sector_offset;
            $this->fatpointer = $this->fat[$this->fatpointer];
            while ($this->fatpointer != OleDocument::ENDOFCHAIN && $bytesread < $bytes) {
                if (($bytes - $bytesread) < $this->sectorsize) {
                    $this->buffer = ($this->readsector)($this->fatpointer);
                    $data .= substr($this->buffer, 0, $bytes - $bytesread);
                    $bytesread = $bytes;
                    break;
                } else {
                    $data .= ($this->readsector)($this->fatpointer);
                    $bytesread += $this->sectorsize;
                }

                $this->fatpointer = $this->fat[$this->fatpointer];
            }
            $this->pos += $bytesread;
        }

        return $data;
    }

    public function write($data)
    {
        // don't need to worry about buffering writes -- OleDocument already buffers all writes
        $byteswritten = $this->root->write($this->entry->getId(), $data, $this->pos);
        $this->doSeek($this->pos + $byteswritten);
        return $byteswritten;
    }

    public function readUint1(): int
    {
        $s = $this->read(1);
        if (!$s) {
            throw new \Exception('unable to read integer value');
        }

        return ord($s[0]);
    }

    public function readUint2(): int
    {
        $s = $this->read(2);
        if (strlen($s) != 2) {
            throw new \Exception('unable to read integer value');
        }

        return ord($s[0]) | (ord($s[1]) << 8);
    }

    public function readUint4(): int
    {
        $s = $this->read(4);
        if (strlen($s) != 4) {
            throw new \Exception('unable to read integer value');
        }

        return ord($s[0]) | (ord($s[1]) << 8) | (ord($s[2]) << 16) | (ord($s[3]) << 24);
    }
}


