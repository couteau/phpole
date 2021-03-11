<?php
namespace Cryptodira\PhpOle;

class OleDirectoryEntry implements OleSerializable
{
    const FORMAT = 'A64v1C1C1V1V1V1H32V1P1P1V1P1';
    const EMPTY = 
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .    // EntryName
        "\000\000" .                                                            // EntryNameLength
        "\000" .                                                                // ObjectType
        "\000" .                                                                // ColorFlag
        "\xFF\xFF\xFF\xFF" .                                                    // LeftSiblingID
        "\xFF\xFF\xFF\xFF" .                                                    // RightSiblingID
        "\xFF\xFF\xFF\xFF" .                                                    // ChildID
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .    // CLSID
        "\000\000\000\000" .                                                    // StateBits
        "\000\000\000\000\000\000\000\000" .                                    // CreationTime
        "\000\000\000\000\000\000\000\000" .                                    // ModifiedTime
        "\xFE\xFF\xFF\xFF" .                                                    // StartingSector
        "\000\000\000\000\000\000\000\000";                                     // StreamSize
    
    const LEFT = -1;
    const RIGHT = 1;
    const DIRECTION = [
        self::LEFT => 'LeftChildID',
        0 => false,
        self::RIGHT => 'RightChildID',
    ];

    private $root;
    private $objectId;
    private $entry;

    public function __construct(OleDocument $root, int $objectId, array $entry)
    {
        $this->root = $root;
        $this->objectId = $objectId;
        $this->entry = $entry;
    }

    public static function new(OleDocument $root, int $objectId, $name, $type, $clsid = null, $color = OleDocument::COLOR_RED)
    {
        // According to MS spec, CLSID must be all 0's for stream objects
        if ($type === OleDocument::StreamObject) {
            $clsid = null;
        }

        $createTime = $type === OleDocument::StreamObject ? 0 : time() * 100000 - Ole::FILETIME_BASE;

        $entry = [
            'EntryName' => substr($name, 0, 32),
            'EntryNameLength' => (strlen($name) + 1) * 2,
            'ObjectType' => $type,
            'ColorFlag'  => $type === OleDocument::RootStorageObject ? OleDocument::COLOR_BLACK : $color,
            'LeftSiblingID' => OleDocument::NOSTREAM,
            'RightSiblingID' => OleDocument::NOSTREAM,
            'ChildID' => OleDocument::NOSTREAM,
            'CLSID' => $clsid ? str_replace(['{','}','-'], '', $clsid) : str_repeat('0', 32),
            'StateBits' => 0,
            'CreationTime' => $createTime,
            'ModifiedTime' => $createTime,
            'StartingSector' => $type === OleDocument::StorageObject ? 0 : OleDocument::ENDOFCHAIN,
            'StreamSize' => 0
        ];

        return new OleDirectoryEntry($root, $objectId, $entry);
    }

    public function getId()
    {
        return $this->objectId;
    }

    public function getName()
    {
        return $this->entry['EntryName'];
    }

    public function getEntryNameLength()
    {
        return $this->entry['EntryNameLength'];
    }

    public function getObjectType()
    {
        return $this->entry['ObjectType'];
    }

    public function getStartingSector()
    {
        return $this->entry['StartingSector'] === OleDocument::NOSTREAM ? null : $this->entry['StartingSector'];
    }

    public function getStreamSize()
    {
        return $this->entry['StreamSize'];
    }

    public function setStreamSize($newsize)
    {
        $this->entry['StreamSize'] = $newsize;
    }

    public function leftChild()
    {
        return $this->entry['LeftSiblingID'] === OleDOcument::NOSTREAM ? null : $this->root[$this->entry['LeftSiblingID']];
    }

    public function rightChild()
    {
        return $this->entry['RightSiblingID'] === OleDOcument::NOSTREAM ? null : $this->root[$this->entry['RightSiblingID']];
    }

    public function isLeaf()
    {
        return $this->entry['LeftSiblingID'] === OleDOcument::NOSTREAM && $this->entry['RightSiblingID'] === OleDOcument::NOSTREAM;
    }

    public function hasChild($direction)
    {
        $this->entry[self::DIRECTION[$direction]] !== OleDocument::NOSTREAM;
    }

    public function getChild($direction)
    {
        $childId = $this->entry[self::DIRECTION[$direction]];
        if ($childId) {
            if ($childId === OleDocument::NOSTREAM) {
                return null;
            } else {
                return $this->root[$childId];
            }
        } else {
            throw new \InvalidArgumentException('direction must be non-zero');
        }
    }

    public function setChild($direction, $child)
    {
        $field = self::DIRECTION[$direction];
        $oldChild = $this->getChild($direction);
        $this->entry[$field] = $child ? $child->getId() : OleDocument::NOSTREAM;
        if ($child) {
            $child->entry['ColorFlag'] = OleDocument::COLOR_RED;
        }
        if ($oldChild && $child) {
            $child->setChild($child->compareTo($oldChild), $oldChild);
        }
        return $oldChild;
    }

    public function getRootEntry() 
    {
        return $this->entry['ChildID'] === OleDocument::NOSTREAM ? null : $this->root[$this->entry['ChildID']];
    }

    public function setRootEntry(OleDirectoryEntry $child) 
    {
       $this->entry['ChildID'] = $child->getId();
       $child->entry['ColorFlag'] = OleDocument::COLOR_BLACK;
       return $this;
    }

    public function serialize(): string
    {
        $entry = $this->entry;
        $entry['EntryName'] = mb_convert_encoding($entry['EntryName'], "UTF-16LE", "UTF-8");
        return pack(self::FORMAT, ...array_values($entry));
    }

    public function compareTo(self $entry)
    {
        if ($this === $entry) {
            return 0;
        } else {
            $r = $this->entry['EntryNameLength'] - $entry->entry['EntryNameLength'];
            if ($r === 0) {
                $r = strcasecmp($this->entry['EntryName'], $entry->entry['EntryName']);
            }
            return $r / abs($r); // -1 for < 0; 1 for > 0
        }
    }

    public function compareToName(string $name)
    {
        if ($this->entry['EntryName'] === $name) {
            return 0;
        } else {
            $r = strlen($this->entry['EntryName']) - strlen($name);
            if ($r === 0) {
                $r = strcasecmp($this->entry['EntryName'], $name);
            }
            return $r / abs($r); // -1 for < 0; 1 for > 0
        }
    }

}