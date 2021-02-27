<?php
namespace Cryptodira\PhpMsOle;

/**
 *
 * @author stuart
 *
 */
class OLEStorage implements \IteratorAggregate, \Countable, \ArrayAccess
{

    /**
     * Array of directory entries in the directory structure
     * Each entry is an array with entries corresponding to the OLEDocument::OLEDirectoryEntryFormat format string
     *
     * @var array
     */
    protected $entries = [];

    private $root;

    private $bs;

    /**
     */
    public function __construct(OLEDocument $root, $streamid = null)
    {
        $this->root = $root;
        if ($this !== $root) {
            $this->readStorageStream($streamid);
        }
    }

    protected function readStorageStream(int $streamid)
    {
        if ($this->root[$streamid]['ObjectType'] != 1) {
            throw new \Exception("Id {$streamid} is not a storage");
        }

        $this->readStorage($this->root[$streamid]['StartingSector']);
    }

    protected function readStorage(int $sector)
    {
        $this->bs = $this->root->getBlocksize();
        $s = $sector;

        while ($s != self::ENDOFCHAIN) {
            $this->readDirectorySector($s);
            $s = $this->root->getNextSector($s);
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
    protected function readDirectorySector($sector)
    {
        $data = $this->root->getSectorData($sector);
        for ($i = 0; $i < $this->bs / 128; $i++) {
            $newentry = unpack(OLEDocument::OLEDirectoryEntryFormat, $data, $i * OLEDocument::OLEDirectoryEntrySize);
            // unpack cuts off the final byte of the final UTF-16LE character if it is null, so we have to add it back on
            if (strlen($newentry['EntryName']) % 2 != 0)
                $newentry['EntryName'] .= chr(0);
            $newentry['EntryName'] = mb_convert_encoding($newentry['EntryName'], "UTF-8", "UTF-16LE");
            $this->entries[] = $newentry;
        }
        return $this->entries;
    }

    public function getIterator()
    {
        return new \ArrayIterator($this->entries);
    }

    public function count()
    {
        return count($this->entries);
    }

    public function offsetGet($offset)
    {
        return $this->entries[$offset];
    }

    public function offsetExists($offset)
    {
        return $offset >= 0 && $offset < count($this->entries);
    }

    public function offsetUnset($offset)
    {
        throw new \BadMethodCallException('Cannot unset OLEDocument directory entries');
    }

    public function offsetSet($offset, $value)
    {
        throw new \BadMethodCallException('Cannot set OLEDocument directory entries');
    }
}