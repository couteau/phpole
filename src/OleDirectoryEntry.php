<?php
namespace Cryptodira\PhpOle;

class OleDirectoryEntry implements OleSerializable
{
    // Unpack() format string for a single directory entry (128 bytes)
    const READFORMAT =
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
        'V1StartingSector/' .    # 74 4 bytes
        'P1StreamSize/';        # 78 8 bytes

    // Pack() format string for a single directory entry (128 bytes)
    const WRITEFORMAT = 'A64v1C1C1V1V1V1H32V1P1P1V1P1';

    // Packed content of an unused directory entry (128 bytes)
    const EMPTY = 
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .    // EntryName
        "\000\000" .                                                            // EntryNameLength
        "\000" .                                                                // ObjectType
        "\000" .                                                                // ColorFlag
        "\xFF\xFF\xFF\xFF" .                                                    // LeftSiblingID = NOSTREAM
        "\xFF\xFF\xFF\xFF" .                                                    // RightSiblingID = NOSTREAM
        "\xFF\xFF\xFF\xFF" .                                                    // ChildID = NOSTREAM
        "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .    // CLSID
        "\000\000\000\000" .                                                    // StateBits
        "\000\000\000\000\000\000\000\000" .                                    // CreationTime
        "\000\000\000\000\000\000\000\000" .                                    // ModifiedTime
        "\xFE\xFF\xFF\xFF" .                                                    // StartingSector = ENDOFCHAIN
        "\000\000\000\000\000\000\000\000";                                     // StreamSize
    
    const LEFT = -1;
    const RIGHT = 1;
    const DIRECTION = [
        self::LEFT => 'LeftSiblingID',
        0 => false,
        self::RIGHT => 'RightSiblingID',
    ];

    private $root;
    private $objectId;
    private $entry;

    public function __construct(OleDocument $root, int $objectId, $entryData, $offset = 0)
    {
        $this->root = $root;
        $this->objectId = $objectId;

        if (is_string($entryData)) {
            $newentry = unpack(self::READFORMAT, $entryData, $offset);
            if ($newentry['ObjectType'] === 0) {
                throw new \InvalidArgumentException('Cannot create a directory entry object from an unused directory entry');
            }

            // unpack cuts off the final byte of the final UTF-16LE character if it is null, so we have to add it back on before converting
            if (strlen($newentry['EntryName']) % 2 != 0) {
                $newentry['EntryName'] .= chr(0);
            }
            $newentry['EntryName'] = mb_convert_encoding($newentry['EntryName'], "UTF-8", "UTF-16LE");

            $this->entry = $newentry;
        } else {
            $this->entry = $entryData;
        }
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

    public function copyTo(OleStorage $dest)
    {
        $destRoot = $dest->getRoot();
        $newentry = new OleDirectoryEntry($destRoot, -1, $this->entry);
        $newentry->entry['LeftSiblingID'] = OleDocument::NOSTREAM;
        $newentry->entry['RightSiblingID'] = OleDocument::NOSTREAM;
        $newentry->entry['ChildID'] = OleDocument::NOSTREAM;
        $newentry->entry['StartingSector'] = $newentry->entry['ObjectType'] === OleDocument::StorageObject ? 0 : OleDocument::ENDOFCHAIN;
        $newentry->entry['StreamSize'] = 0;
        $objectId = $destRoot->addEntry($newentry, false);
        $newentry->objectId = $objectId;
        $dest->insertEntry($newentry);

        if ($this->entry['ObjectType'] === OleDocument::StreamObject) {
            $data = $this->root->getStreamData($this->objectId);
            $destRoot->write($objectId, $data);
        } else {
            $storage = $this->root->getObject($this);
            $newstorage = $destRoot->getObject($newentry);

            // copy by walking the tree using the Preorder algorithm
            // to help keep the tree balanced in the destination storage
            // (assuming source is balanced) since we don't implement a r-b tree
            $storage->foreach(function ($entry) use ($newstorage) {
                $entry->copyTo($newstorage);
            }, OleStorage::PREORDER);
        }
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
        return $this->entry['StartingSector'];
    }

    public function &_getStartingSector()
    {
        $result =& $this->entry['StartingSector'];
        return $result;
    }

    public function setStartingSector($sector)
    {
        $this->entry['StartingSector'] = $sector;
    }

    public function getStreamSize()
    {
        return $this->entry['StreamSize'];
    }

    public function &_getStreamSize()
    {
        $result =& $this->entry['StreamSize'];
        return $result;
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
        return $this->entry[self::DIRECTION[$direction]] !== OleDocument::NOSTREAM;
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
        return pack(self::WRITEFORMAT, ...array_values($entry));
    }

    // Microsoft OLE spec says entries are sorted by name length and then by a case insenstivie comparison of the name
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