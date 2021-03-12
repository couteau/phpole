<?php
namespace Cryptodira\PhpOle;

class OleHeader implements OleSerializable
{
    // Magic number identifying file as an Ole compound file
    const MAGIC = 'd0cf11e0a1b11ae1';

    // Unpack() format string for the Ole Header
    const READFORMAT = 
            'H16Signature/' .           # 00 8 bytes
            'H32CLSID/' .               # 08 16 bytes
            'v1MinorVersion/' .         # 18 2 bytes
            'v1MajorVersion/' .         # 1A 2 bytes
            'v1ByteOrder/' .            # 1C 2 bytes
            'v1SectorShift/' .          # 1E 2 bytes
            'v1MiniSectorShift/' .      # 20 2 bytes
            'Z6Reserved1/' .            # 22 6 bytes
            'V1DirectorySectors/' .     # 28 4 bytes
            'V1FATSectors/' .           # 2C 4 bytes
            'V1FirstDirectorySector/' . # 30 4 bytes
            'V1TransactionSignature/' . # 34 4 bytes
            'V1MiniStreamCutoff/' .     # 38 4 bytes
            'V1FirstMiniFATSector/' .   # 3C 4 bytes
            'V1MiniFATSectorCount/' .   # 40 4 bytes
            'V1FirstDIFATSector/' .     # 44 4 bytes
            'V1DIFATSectorCount/';      # 48 4 bytes
            
    
    private $root;

    private $dirty;

    private $header;

    public function __construct(OleDocument $root, string $headerData)
    {
        $this->root = $root;
        $this->header = unpack(self::READFORMAT, $headerData);

        if ($this->header['Signature'] != self::MAGIC) {
            throw new \Exception("Stream is not an Ole file");
        }
    }

    public static function new($root, $version = 0x0003)
    {   
        $data = "\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" .                                // Magic
            "\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000\000" .    // CLSID
            "\x3E\x00" .                                                            // MinorVersion
            pack('v1', $version) .	                                                // MajorVersion
            "\xFE\xFF" .	                                                        // ByteOrder
            ($version === 0x0003 ? "\x09\x00" : "\x0C\x00") .	                    // SectorShift
            "\x06\x00" .	                                                        // MiniSectorShift
            "\000\000\000\000\000\000" .                        	                // Reserved1
            "\000\000\000\000" .                                                    // DirectorySectors
            "\000\000\000\000" .            	                                    // FATSectors
            pack('V1', OleDocument::ENDOFCHAIN) .           	                    // FirstDirectorySector
            "\000\000\000\000" .            	                                    // TransactionSignature
            "\x00\x10\x00\x00" .	                                                // MiniStreamCutoff
            pack('V1', OleDocument::ENDOFCHAIN) .                                   // FirstMiniFATSector
            "\000\000\000\000" .            	                                    // MiniFATSectorCount
            pack('V1', OleDocument::ENDOFCHAIN) .             	                    // FirstDIFATSector
            "\000\000\000\000";                                                      // DIFATSectorCount

        $header = new self($root, $data);
        $header->dirty = true;
        return $header;
    }

    public function isDirty()
    {
        return $this->dirty;
    }

    public function makeDirty()
    {
        $this->dirty = true;
    }

    public function makeClean()
    {
        $this->dirty = false;
    }

    public function getVersion()
    {
        return $this->header['MajorVersion'];
    }

    public function getSectorShift()
    {
        return $this->header['SectorShift'];
    }

    public function getMiniSectorShift()
    {
        return $this->header['MiniSectorShift'];
    }

    public function getMiniStreamCutoff()
    {
        return $this->header['MiniStreamCutoff'];
    }

    public function getFirstDIFATSector()
    {
        return $this->header['FirstDIFATSector'];
    }

    public function setFirstDIFATSector($sector)
    {
        $this->header['FirstDIFATSector'] = $sector;
        $this->dirty = true;
    }

    public function incDIFATSectorCount()
    {
        $this->header['DIFATSectorCount'] = $this->header['DIFATSectorCount'] + 1;
        $this->dirty = true;
    }

    public function incFATSectorCount()
    {
        $this->header['FATSectors'] = $this->header['FATSectors'] + 1;
        $this->dirty = true;
    }

    public function getFirstDirectorySector()
    {
        return $this->header['FirstDirectorySector'];
    }

    public function setFirstDirectorySector($sector)
    {
        $this->header['FirstDirectorySector'] = $sector;
        if ($this->header['MajorVersion'] === 0x0004) {
            $this->header['DirectorySectors'] = 1;
        }
        $this->dirty = true;
    }

    public function getFirstMiniFATSector()
    {
        return $this->header['FirstMiniFATSector'];
    }

    public function setFirstMiniFATSector($sector)
    {
        $this->header['FirstMiniFATSector'] = $sector;
        $this->header['MiniFATSectorCount'] = 1;
        $this->dirty = true;
    }

    public function serialize(): string
    {
        $headerValues = array_values($this->header);

        $fmt = preg_replace('/([A-Za-z][0-9]+)[A-Z]\w+(\/|$)/', '$1', self::READFORMAT);
        return pack($fmt, ...$headerValues);
    }
}