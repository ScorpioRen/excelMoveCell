<?php
/**
 * Excel 单元格移动
 *
 * PHP version 7.0
 * Date: 16-12-22
 * Time: 上午11:27
 *
 * @category PHP
 * @package  Platform\Module\Single\ExcelModelBundle\Helper
 * @author   viron <viron@foxmail.com>
 * @license  http://www.acupwater.com BSD listen
 * @link     http://www.acupwater.com
 */
class ExcelHelper
{

    //字母索引
    protected static $ALPHABET = [0,
        'A','B','C','D','E','F','G','H','I','G','K','L','M',
        'N','O','P','Q','R','S','T','U','V','W','X','Y','Z'
    ];

    /**
     * 相对位移
     *
     * @param string  $cell       单元格 C1
     * @param integer $horizontal 横向了列
     * @param integer $vertical   纵向行
     *
     * @return string
     */
    public static function offset($cell, $horizontal = 0, $vertical = 0)
    {
        $cellName = self::getCellSplit($cell);
        $columnName = self::moveColumn($cellName['column'], $horizontal);
        $row = self::moveRow($cellName['row'], $vertical);
        return $columnName.$row;
    }

    /**
     * 上下移动
     *
     * @param integer $row  行
     * @param integer $step 移动个数
     *
     * @return integer
     */
    protected static function moveRow($row, $step)
    {
        $row = (int)$row+(int)$step;
        if (0 >= $row) {
            return 1;
        }
        return $row;
    }

    /**
     * 拆分单元格名称
     *
     * @param string $cell 单元格名称
     *
     * @return array|bool
     */
    protected static function getCellSplit($cell)
    {
        if (preg_match('/^([a-zA-Z]+)([0-9]+)$/', $cell, $res)) {
            return ['column'=>strtoupper($res[1]),'row'=>$res[2]];
        }
        return false;
    }

    /**
     * 左右移动
     *
     * @param string  $column 列明
     * @param integer $step   步长
     *
     * @return string
     */
    protected static function moveColumn($column, $step)
    {
        $cNumber = self::alphabetToNumber($column);
        return self::numberToAlphabet($cNumber+(int)$step);
    }

    /**
     * 数字转换字母
     *
     * @param integer $number 数字
     *
     * @return mixed|string
     */
    protected static function numberToAlphabet($number)
    {

        $index = $number%27==0?1:$number%27;
        $res = self::$ALPHABET[$index];
        $tmp = intval($number/27);
        if ($tmp > 0) {
            $res = self::numberToAlphabet($tmp).$res;
        }
        return $res;
    }

    /**
     * 字符转数字
     *
     * @param string $alphabet 字符
     *
     * @return int
     */
    protected static function alphabetToNumber($alphabet)
    {
        $alphabetFlip = array_flip(self::$ALPHABET);
        $cLen = strlen($alphabet);
        $res = 0;
        for ($i=$cLen-1; $i>=0; $i--) {
            $res += $alphabetFlip[$alphabet[$i]] * pow(26, $i);
        }
        return $res;
    }
}
