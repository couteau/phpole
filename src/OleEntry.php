<?php
namespace Cryptodira\PhpOle;

/**
 *
 * @author stuart
 *
 */
class OleEntry
{

    protected $root;

    protected $entry;

    public function __construct(OleDocument $root, $stream)
    {
        $this->root = $root;
        if (is_null($stream)) {
            $stream = $root->getDocumentStream();
            $filespec = $root[$stream];
        } elseif (is_string($stream)) {
            if (!$stream = $root->FindStreamByName($stream)) {
                throw new \Exception("Stream {$stream} not found");
            }
            $filespec = $root[$stream];
        } elseif (is_int($stream)) {
            $filespec = $root[$stream];
        } elseif (is_array($stream)) {
            $filespec = $stream;
        } else {
            throw new \Exception("Invalid stream {$stream}");
        }

        if (static::class != OleDocument::TYPE_MAP[$filespec['ObjectType']]) {
            throw new \Exception($filespec['EntryName'] . " is not the correct type for " . static::class);
        }

        $this->entry = $filespec;
    }
}

