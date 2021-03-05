<?php
namespace Cryptodira\PhpOle;

/**
 *
 * @author stuart
 *
 */
class OleStorage extends OleEntry implements \IteratorAggregate, \Countable, \ArrayAccess
{

    /**
     * Array of directory entries in the directory structure
     * Each entry is an array with entries corresponding to the OleDocument::OleDirectoryEntryFormat format string
     *
     * @var array
     */
    protected $entries = [];

    protected $nameMap = [];

    /**
     */
    public function __construct(OleDocument $root, $stream = null)
    {
        parent::__construct($root, $stream);
        $this->readStorageStream();
    }

    private function visitNode($streamid, $callback)
    {
        $entry = $this->root[$streamid];

        if ($entry['LeftSiblingID'] != OleDocument::FREESECT) {
            $this->visitNode($entry['LeftSiblingID'], $callback);
        }
        $callback($entry, $streamid);
        if ($entry['RightSiblingID'] != OleDocument::FREESECT) {
            $this->visitNode($entry['RightSiblingID'], $callback);
        }
    }

    protected function readStorageStream()
    {
        if ($this->entry['ChildID'] != OleDocument::FREESECT) {
            $this->visitNode($this->entry['ChildID'],
                    function ($child, $streamid) {
                        $this->entries[$streamid] = $child;
                        $this->nameMap[$child['EntryName']] = $streamid;
                    });
        }
    }

    public function foreach(\Closure $callback)
    {
        $this->visitNode($this->entry['ChildID'], $callback);
    }

    /**
     * Return the streamid for the stream with the passed stream name
     *
     * @param string $streamName
     * @return number|boolean
     */
    public function findStreamByName($streamName)
    {
        return $this->nameMap[$streamName] ?? false;
    }

    public function getStreamById($streamId)
    {
        return $this->entries[$streamId] ?? false;
    }

    public function getData($streamId)
    {
        if (array_key_exists($streamId, $this->entries)) {
            return $this->root->getData($streamId);
        } else {
            return null;
        }
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
        return array_key_exists($offset, $this->entries);
    }

    public function offsetUnset($offset)
    {
        throw new \BadMethodCallException('OleDocument directory entries are readonly');
    }

    public function offsetSet($offset, $value)
    {
        throw new \BadMethodCallException('OleDocument directory entries are readonly');
    }
}