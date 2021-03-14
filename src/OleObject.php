<?php
namespace Cryptodira\PhpOle;

/**
 *
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
class OleObject
{

    protected $root;

    /**
     * The Directory Entry within the root OLE file for this object
     *
     * @var OleDirectoryEntry
     */
    protected $entry;

    public function __construct(OleDocument $root, OleDirectoryEntry $entry = null)
    {
        $this->root = $root;
        if (is_null($entry)) {
            $stream = $root->getDocumentStream();
            $entry = $root[$stream];
        }

        if (static::class != OleDocument::TYPE_MAP[$entry->getObjectType()]) {
            throw new \Exception($entry->getName() . " is not the correct type for " . static::class);
        }

        $this->entry = $entry;
    }

    public function getEntry()
    {
        return $this->entry;
    }

    public function getRoot()
    {
        return $this->root;
    }

    public function copyTo(OleStorage $dest)
    {
        $this->entry->copyTo($dest);
    }
}
