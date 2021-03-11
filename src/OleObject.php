<?php
namespace Cryptodira\PhpOle;

/**
 *
 * @author stuart
 *
 */
class OleObject
{

    protected $root;

    protected $entry;

    public function __construct(OleDocument $root, $stream = null)
    {
        $this->root = $root;
        if (is_null($stream)) {
            $stream = $root->getDocumentStream();
            $filespec = $root[$stream];
        } elseif (is_string($stream)) {
            if (!$stream = $root->findEntryByName($stream)) {
                throw new \Exception("Stream {$stream} not found");
            }
            $filespec = $root[$stream];
        } elseif (is_int($stream)) {
            $filespec = $root[$stream];
        } elseif ($stream instanceof OleDirectoryEntry) {
            $filespec = $stream;
        } else {
            throw new \Exception("Invalid stream {$stream}");
        }

        if (static::class != OleDocument::TYPE_MAP[$filespec->getObjectType()]) {
            throw new \Exception($filespec->getName() . " is not the correct type for " . static::class);
        }

        $this->entry = $filespec;
    }

    public function getEntry()
    {
        return $this->entry;
    }
}

