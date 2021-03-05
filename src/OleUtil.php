<?php

$CODEPAGE = 1252;

const mb_encodings = array(
    1200 => 'UTF-16LE',
    1201 => 'UTF-16BE',
    1251 => 'CP1251',
    1252 => 'CP1252'
);

const iconv_encodings = array(
    10000 => 'macintosh',
    10004 => 'macarabic',
    10005 => 'machebrew',
    10006 => 'macgreek',
    10007 => 'maccyrillic',
    10010 => 'macromania',
    10017 => 'macukraine',
    10021 => 'macthai',
    10029 => 'maccentraleurope',
    10079 => 'maciceland',
    10081 => 'macturkish',
    10082 => 'maccroatia',
    20127 => 'ASCII',
    28591 => 'ISO-8859-1',
    28592 => 'ISO-8859-2',
    28593 => 'ISO-8859-3',
    28594 => 'ISO-8859-4',
    28595 => 'ISO-8859-5',
    28596 => 'ISO-8859-6',
    28597 => 'ISO-8859-7',
    28598 => 'ISO-8859-8',
    28599 => 'ISO-8859-9',
    50220 => 'ISO-2022',
    50221 => 'ISO-2022-1',
    50222 => 'ISO-2022-2',
    50225 => 'ISO-2022-KR',
    50227 => 'ISO-2022-CN',
    50229 => 'ISO-2022-CN-EXT',
    65000 => 'UTF-8',
    65001 => 'UTF-7'
);

function ror($c, $bits)
{
    $b = str_pad(decbin(ord($c) & 0xFF), 8, '0', STR_PAD_LEFT);
    $b = substr($b, 8 - $bits, $bits) . substr($b, 0, 8 - $bits);
    return chr(bindec($b));
}

function ole_convert_encoding($str, $codepage = null)
{
    global $CODEPAGE;

    if (is_null($codepage))
        $codepage = $CODEPAGE;

    if (isset(mb_encodings[$codepage]))
        $out = mb_convert_encoding($str, 'UTF-8', mb_encodings[$codepage]);
    elseif (isset(iconv_encodings[$codepage])) {
        $out = iconv(iconv_encodings[$codepage], 'UTF-8//TRANSLIT', $str);
        if (!$out)
            $out = iconv(iconv_encodings[$codepage], 'UTF-8//IGNORE', $str);
    } elseif ($codepage >= 437 && $codepage <= 1258) {
        $cp = 'CP' . $codepage;
        $out = iconv($cp, 'UTF-8//TRANSLIT', $str);
        if (!$out)
            $out = iconv($cp, 'UTF-8//IGNORE', $str);
        if (!$out)
            $out = $str;
    } else
        $out = $str;

    return rtrim($out, "\0");
}

function ole_read_unicode_string($buffer, $offset, &$bytesread)
{
    $size = unpack('V1', $buffer, $offset)[1];
    if ($size == 0)
        return '';

    $str = substr($buffer, $offset + 4, $size * 2);
    $bytesread = 4 + $size * 2;
    return mb_convert_encoding($str, 'UTF-8', 'UTF-16LE');
}

function ole_read_codepage_string($buffer, $offset, &$bytesread, $codepage = null)
{
    $size = unpack('V1', $buffer, $offset)[1];
    if ($size == 0)
        return '';

    $str = substr($buffer, $offset + 4, $size);
    $bytesread = 4 + $size;
    return ole_convert_encoding($str, $codepage);
}