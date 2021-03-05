<?php

namespace Cryptodira\PhpOle\Tests;

use PHPUnit\Framework\TestCase;
use Cryptodira\PhpOle\OLEDocument;
use Cryptodira\PhpOle\OLEStorage;

/**
 * OLEDocument test case.
 */
class OleDocumentTest extends TestCase
{

    /**
     *
     * @var OLEDocument
     */
    private $oLEDocument;

    /**
     * Prepares the environment before running a test.
     */
    protected function setUp(): void
    {
        parent::setUp();

        $this->oLEDocument = new OLEDocument(__DIR__ . '/../samples/DOCS-#602483-v1-PPP_Compliance_Survey.doc');
    }

    /**
     * Cleans up the environment after running a test.
     */
    protected function tearDown(): void
    {
        $this->oLEDocument = null;

        parent::tearDown();
    }

    /**
     * Tests OLEDocument->close()
     */
    public function testClose()
    {
        $this->oLEDocument->close();
    }

    /**
     * Tests OLEDocument->getRootStorage()
     */
    public function testGetRootStorage()
    {
        $storage = $this->oLEDocument->getRootStorage();
        $this->assertIsObject($storage);
        $this->assertInstanceOf(OLEStorage::class, $storage);
    }

    /**
     * Tests OLEDocument->getRootStorageCount()
     */
    public function testGetRootStorageCount()
    {
        // TODO Auto-generated OleDocumentTest->testGetRootStorageCount()
        $this->markTestIncomplete("getRootStorageCount test not implemented");

        $this->oLEDocument->getRootStorageCount(/* parameters */);
    }

    /**
     * Tests OLEDocument->getBlocksize()
     */
    public function testGetBlocksize()
    {
        $blocksize = $this->oLEDocument->getBlocksize();
        $this->assertEquals(512, $blocksize);
    }

    /**
     * Tests OLEDocument->getNextBlock()
     */
    public function testGetNextBlock()
    {
        // TODO Auto-generated OleDocumentTest->testGetNextBlock()
        $this->markTestIncomplete("getNextBlock test not implemented");

        $this->oLEDocument->getNextBlock(0);
    }

    /**
     * Tests OLEDocument->getBlockData()
     */
    public function testGetBlockData()
    {
        // TODO Auto-generated OleDocumentTest->testGetBlockData()
        $this->markTestIncomplete("getBlockData test not implemented");

        $this->oLEDocument->getBlockData(0);
    }

    /**
     * Tests OLEDocument->getMiniBlockData()
     */
    public function testGetMiniBlockData()
    {
        // TODO Auto-generated OleDocumentTest->testGetMiniBlockData()
        $this->markTestIncomplete("getMiniBlockData test not implemented");

        $this->oLEDocument->getMiniBlockData(0);
    }

    /**
     * Tests OLEDocument->read()
     */
    public function testRead()
    {
        // TODO Auto-generated OleDocumentTest->testRead()
        $this->markTestIncomplete("read test not implemented");

        $this->oLEDocument->read(1);
    }

    /**
     * Tests OLEDocument->getData()
     */
    public function testGetData()
    {
        // TODO Auto-generated OleDocumentTest->testGetData()
        $this->markTestIncomplete("getData test not implemented");

        $this->oLEDocument->getData(1);
    }

    /**
     * Tests OLEDocument->getDocumentStream()
     */
    public function testGetDocumentStream()
    {
        // TODO Auto-generated OleDocumentTest->testGetDocumentStream()
        $this->markTestIncomplete("getDocumentStream test not implemented");

        $this->oLEDocument->getDocumentStream();
    }

    /**
     * Tests OLEDocument->getStream()
     */
    public function testGetStream()
    {
        // TODO Auto-generated OleDocumentTest->testGetStream()
        $this->markTestIncomplete("getStream test not implemented");

        $this->oLEDocument->getStream(1);
    }

    /**
     * Tests OLEDocument->findStreamByName()
     */
    public function testFindStreamByName()
    {
        // TODO Auto-generated OleDocumentTest->testFindStreamByName()
        $this->markTestIncomplete("findStreamByName test not implemented");

        $this->oLEDocument->findStreamByName('DocumentProperties');
    }

    /**
     * Tests OLEDocument->CreateFromStream()
     */
    public function testCreateFromStream()
    {
        // TODO Auto-generated OleDocumentTest->testCreateFromStream()
        $this->markTestIncomplete("CreateFromStream test not implemented");

        $this->oLEDocument->CreateFromStream(/* parameters */);
    }

    /**
     * Tests OLEDocument->CreateFromFile()
     */
    public function testCreateFromFile()
    {
        // TODO Auto-generated OleDocumentTest->testCreateFromFile()
        $this->markTestIncomplete("CreateFromFile test not implemented");

        $this->oLEDocument->CreateFromFile(/* parameters */);
    }

    /**
     * Tests OLEDocument->CreateFromString()
     */
    public function testCreateFromString()
    {
        // TODO Auto-generated OleDocumentTest->testCreateFromString()
        $this->markTestIncomplete("CreateFromString test not implemented");

        $this->oLEDocument->CreateFromString(/* parameters */);
    }

    /**
     * Tests OLEDocument->getIterator()
     */
    public function testGetIterator()
    {
        // TODO Auto-generated OleDocumentTest->testGetIterator()
        $this->markTestIncomplete("getIterator test not implemented");

        $this->oLEDocument->getIterator(/* parameters */);
    }

    /**
     * Tests OLEDocument->count()
     */
    public function testCount()
    {
        // TODO Auto-generated OleDocumentTest->testCount()
        $this->markTestIncomplete("count test not implemented");

        $this->oLEDocument->count(/* parameters */);
    }

    /**
     * Tests OLEDocument->offsetGet()
     */
    public function testOffsetGet()
    {
        // TODO Auto-generated OleDocumentTest->testOffsetGet()
        $this->markTestIncomplete("offsetGet test not implemented");

        $this->oLEDocument->offsetGet(/* parameters */);
    }

    /**
     * Tests OLEDocument->offsetExists()
     */
    public function testOffsetExists()
    {
        // TODO Auto-generated OleDocumentTest->testOffsetExists()
        $this->markTestIncomplete("offsetExists test not implemented");

        $this->oLEDocument->offsetExists(/* parameters */);
    }

    /**
     * Tests OLEDocument->offsetUnset()
     */
    public function testOffsetUnset()
    {
        // TODO Auto-generated OleDocumentTest->testOffsetUnset()
        $this->markTestIncomplete("offsetUnset test not implemented");

        $this->oLEDocument->offsetUnset(/* parameters */);
    }

    /**
     * Tests OLEDocument->offsetSet()
     */
    public function testOffsetSet()
    {
        // TODO Auto-generated OleDocumentTest->testOffsetSet()
        $this->markTestIncomplete("offsetSet test not implemented");

        $this->oLEDocument->offsetSet(/* parameters */);
    }
}
