<?php
namespace Cryptodira\PhpOle;

/**
 *
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
class OleStorage extends OleObject implements \IteratorAggregate, \Countable, \ArrayAccess
{
    const INORDER = 0;
    const PREORDER = -1;
    const POSTORDER = 1;
    /**
     * Root entry of red-black directory tree
     *
     * @var OleDirectoryEntry
     */
    private $rootEntry;

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
    public function __construct(OleDocument $root, OleDirectoryEntry $entry = null)
    {
        if ($entry === null) {
            $entry = $root[0];
        }
        parent::__construct($root, $entry);
        $this->readStorage();
    }

    private function walkTree($entry, $callback, $order = self::INORDER)
    {
        if ($order >= self::INORDER && $entry->leftChild()) {
            $this->walkTree($entry->leftChild(), $callback, $order);
            if ($order === self::POSTORDER && $entry->rightChild()) {
                $this->walkTree($entry->rightChild(), $callback, $order);
            }
        }

        $callback($entry);

        if ($order <= self::INORDER) {
            if ($order === self::PREORDER && $entry->leftChild()) {
                $this->walkTree($entry->leftChild(), $callback, $order);
            }
            if ($entry->rightChild()) {
                $this->walkTree($entry->rightChild(), $callback, $order);
            }        
        }
    }

    protected function readStorage()
    {
        $this->rootEntry = $this->entry->getRootEntry();
        if ($this->rootEntry) {
            $this->walkTree($this->rootEntry,
                    function ($child) {
                        $this->entries[$child->getId()] = $child;
                        $this->nameMap[$child->getName()] = $child;
                    });
        }
    }

    public function foreach(\Closure $callback, $order = self::INORDER)
    {
        $this->walkTree($this->rootEntry, $callback, $order);
    }

    /**
     * Return the streamid for the stream with the passed stream name
     *
     * @param string $streamName
     * @return number|boolean
     */
    public function findEntryByName($entryName)
    {
        // could use a b-tree search algorithm for this, but most storages 
        // have fewer than 100 entries, so that's probably overkill
        return $this->nameMap[$entryName] ?? false;
    }

    public function getDirectoryEntryById($entryId)
    {
        return $this->entries[$entryId] ?? false;
    }

    public function getObject($entry)
    {
        return $this->root->getObject($entry);
    }

    public function getStreamData($streamId)
    {
        if ($streamId instanceof OleDirectoryEntry) {
            $streamId = $streamId->getId();
        }

        if (array_key_exists($streamId, $this->entries)) {
            return $this->root->getStreamData($streamId);
        } else {
            return null;
        }
    }
    
    /**
     * Recursive search for the parent node at which to insert $entry
     *
     * @param OleDirectoryEntry $start - the position to search from 
     * @param OleDirectoryEntry $entry - the entry to search for
     * @return OleDirectoryEntry - the parent below which the entry should be inserted
     * @throws \InvalidArgumentException
     */
    private function findInsertionPoint($start, $entry)
    {
        if ($start->isLeaf($start)) {
            return $start;
        }
        $direction = $start->compareTo($entry);
        if ($direction === 0 || !$start->hasChild($direction)) {
            return $start;
        }

        return $this->findInsertionPoint($start->getChild($direction), $entry);
    }

    /**
     * Insert a directory entry
     *
     * @param  OleDirectoryEntry $entry
     * @return $this
     */
    public function insertEntry($entry)
    {
        // If root is null, set new entry as root
        if ($this->rootEntry === null) {
            $this->entry->setRootEntry($entry);
        } else {
            $insertNode = $this->findInsertionPoint($this->rootEntry, $entry);
            $insertNode->setChild($insertNode->compareTo($entry), $entry);
        }

        // reload the entries and nameMap arrays
        $this->readStorage();

        return $this;
    }

    public function getIterator()
    {
        return new \ArrayIterator($this->entries);
    }

    public function count()
    {
        return count($this->entries);
    }

    public function &offsetGet($offset)
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