<?php
namespace Cryptodira\PhpOle;

/**
 * Class for reading/manipulating property bag streams in an Ole file
 *
 * @author Stuart C. Naifeh <stuart@cryptodira.org>
 *
 */
abstract class OlePropertyBag implements \ArrayAccess
{

    // Unpack() format string for the individual property header
    const PropertyFormat = 'V1Identifier/V1Offset';

    // Unpack() format string for the property value
    const PropertyValueFormat = 'v1Type/v1Padding';

    // Unpack() format strings for individual property types
    private static $ValueFormats = [
        VT_EMPTY => '',
        VT_NULL => '',
        VT_I2 => 'C1Value1/c1Value2/v1Padding',
        VT_I4 => 'C3Value/c1Value4',
        VT_R4 => 'g1Value',
        VT_R8 => 'e1Value',
        VT_CY => 'C7Value/c1Value8',
        VT_DATE => 'e1Value',
        VT_BSTR => '',
        VT_ERROR => 'V1Value',
        VT_BOOL => 'C1Value/C3',
        VT_DECIMAL => 'v1Reserved/c1Scale/c1Sign/V1Hi32/P1Lo64',
        VT_I1 => 'c1Value/C3',
        VT_UI1 => 'C1Value/C3Padding',
        VT_UI2 => 'v1Value/v1Padding',
        VT_UI4 => 'V1Value',
        VT_I8 => 'C7Value/c1Value8',
        VT_UI8 => 'P1Value',
        VT_INT => 'C3Value/c1Value4',
        VT_UINT => 'V1Value',
        VT_LPSTR => '',
        VT_LPWSTR => 'V1Size',
        VT_FILETIME => 'V1dwLowDateTime/V1dwHighDateTime',
        // VT_BLOB => '',
        // VT_STREAM => '',
        // VT_STORAGE => '',
        // VT_STREAMED_Object => '',
        // VT_STORED_Object => '',
        // VT_BLOB_Object => '',
        // VT_CF => '',
        VT_CLSID => 'V1a/v2b/C8c',
        // VT_VERSIONED_STREAM => '',
        // VT_VECTOR => '',
        // VT_ARRAY => '',
    ];

    /**
     * Header of the preoperty bag
     *
     * @var array
     */
    protected $header;

    protected $headerSize;

    /**
     * The property name => value dictionary for all properties in the stream
     *
     * @var array
     */
    protected $properties;

    /**
     * The codepage in which string property values are encoded
     *
     * @var int
     */
    protected $codepage;

    /**
     *
     * {@inheritdoc}
     * @see \ArrayAccess::offsetExists()
     */
    public function offsetExists($offset)
    {
        return array_key_exists($offset, $this->properties);
    }

    /**
     *
     * {@inheritdoc}
     * @see \ArrayAccess::offsetGet()
     */
    public function offsetGet($offset)
    {
        return $this->properties[$offset];
    }

    /**
     *
     * {@inheritdoc}
     * @see \ArrayAccess::offsetSet()
     */
    public function offsetSet($offset, $value)
    {
        throw new \BadMethodCallException('OlePropertyBag: Properties are read-only');
    }

    /**
     *
     * {@inheritdoc}
     * @see \ArrayAccess::offsetUnset()
     */
    public function offsetUnset($offset)
    {
        throw new \BadMethodCallException('OlePropertyBag: Properties are read-only');
    }

    /**
     * Create a new property bag and read the property sets/properties from the passed binary data
     *
     * @param string $data
     */
    public function __construct($data)
    {
        $this->headerSize = $this->readHeader($data);
        $this->properties = $this->readProperties($data);
    }

    abstract protected function readHeader($data);

    abstract protected function readProperties($data);

    /**
     * Read a property value from the stream starting at offset, and update $offset to end of value data
     *
     * @param string $data
     * @param int $offset
     * @param int $type
     * @return mixed
     */
    protected function readPropertyValue($data, &$offset, $type = null)
    {
        if (!$type || $type == VT_VARIANT) {
            $type = unpack(self::PropertyValueFormat, $data, $offset)['Type'];
            $offset += 4;
        }

        if (isset(self::$ValueFormats[$type]) && !empty(self::$ValueFormats[$type]))
            $value = unpack(self::$ValueFormats[$type], $data, $offset);
        else
            $value = null;

        switch ($type) {
            case VT_EMPTY:
            case VT_NULL:
                $value = null;
                break;
            case VT_I2:
                $value = ($value['Value2'] << 8) | $value['Value1'];
                $offset += 4;
                break;
            case VT_I4:
            case VT_INT:
                $value = ($value['Value4'] << 24) | ($value['Value3'] << 16) | ($value['Value2'] << 8) | $value['Value1'];
                $offset += 4;
                break;
            case VT_DATE:
                $value = new \DateTime();
                $value->setTimestamp(MSDATETIME_BASE + $value['Value'] * 86400);
                $offset += 8;
                break;
            case VT_BOOL:
                $value = $value['Value'] ? true : false;
                $offset += 4;
                break;
            case VT_DECIMAL:
                $v = bcdiv(bcadd(bcmul($value['Hi32'], bcpow(2, 64)), $value['Lo64']), bcpow(10, $value['Scale']),
                        $value['Scale']);
                if ($value['Sign'] == 0x80)
                    $value = bcmul(-1, $v, $value['Scale']);
                else
                    $value = $v;
                $offset += 16;
                break;
            case VT_CY:
            case VT_I8:
                $value = ($value['Value8'] << 56) | ($value['Value7'] << 48) | ($value['Value6'] << 40) |
                        ($value['Value5'] << 32) | ($value['Value4'] << 24) | ($value['Value3'] << 16) |
                        ($value['Value2'] << 8) | $value['Value1'];
                if ($type == VT_CY)
                    $value /= 10000;
                $offset += 8;
                break;
            case VT_BSTR:
            case VT_LPSTR:
                $bytesread = 0;
                $value = ole_read_codepage_string($data, $offset, $bytesread, $this->codepage);
                // $offset += $bytesread + ((4 - ($value['Size'] % 4) % 4));
                $offset += $bytesread; // Padding to 4 byte boundary doesn't seem to be happening notwithstanding the spec
                break;
            case VT_LPWSTR:
                $bytesread = 0;
                $value = ole_read_unicode_string($data, $offset, $bytesread);
                $offset += $bytesread;
                break;
            case VT_FILETIME:
                $ft = ($value['dwHighDateTime'] << 32) | $value['dwLowDateTime'];
                if ($ft == 0)
                    $value = null;
                else {
                    $value = new \DateTime();
                    $value->setTimestamp(FILETIME_BASE + $ft / 10000000);
                }
                break;
            case VT_CLSID:
                $value = vsprintf('{%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}', $value);
                $offset += 16;
                break;
            // Not supporting these types yet
            case VT_BLOB:
            case VT_STREAM:
            case VT_STORAGE:
            case VT_STREAMED_Object:
            case VT_STORED_Object:
            case VT_BLOB_Object:
            case VT_CF:
            case VT_VERSIONED_STREAM:
                $value = null;
                break;
            default:
                if ($type & VT_VECTOR) {
                    $size = unpack('V1', $data, $offset)[1];
                    $value = array();
                    $offset += 4;
                    for ($i = 0; $i < $size; $i++)
                        $value[] = $this->readPropertyValue($data, $offset, $type & 0xFFF);
                } else if ($type & VT_ARRAY)
                    $value = null; // Not supporting Arrays for now
                else
                    $value = $value['Value'];
        }

        return $value;
    }

    public function getProperties()
    {
        return $this->properties;
    }
}

