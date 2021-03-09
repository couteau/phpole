<?php

namespace Cryptodira\PhpOle\Tests;

use PHPUnit\Framework\TestCase;
use Cryptodira\PhpOle\OleDocument;
use Cryptodira\PhpOle\OleStorage;

/**
 * OLEDocument test case.
 */
class OleDocumentTest extends TestCase
{

    /**
     *
     * @var OLEDocument
     */
    private $oleDocument;

    /**
     * Prepares the environment before running a test.
     */
    protected function setUp(): void
    {
        parent::setUp();

        $this->oleDocument = new OleDocument(__DIR__ . '/../samples/TestWordDoc.doc');
    }

    /**
     * Cleans up the environment after running a test.
     */
    protected function tearDown(): void
    {
        $this->oleDocument = null;

        parent::tearDown();
    }

    /**
     * Tests OLEDocument->close()
     */
    public function testClose()
    {
        $this->oleDocument->close();
    }

    /**
     * Tests OLEDocument->getRootStorage()
     */
    public function testGetRootStorage()
    {
        $storage = $this->oleDocument->getRootStorage();
        $this->assertIsObject($storage);
        $this->assertInstanceOf(OleStorage::class, $storage);
    }

    public function testFormat()
    {
        $this->oleDocument = new OleDocument();
        $fn = tempnam(sys_get_temp_dir(), 'ole');
        $this->oleDocument->new($fn);
        $this->oleDocument->close();

        $this->oleDocument = new OleDocument($fn);
        $this->assertEquals(1, count($this->oleDocument));
    }
}
