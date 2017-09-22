<?php
/**
 * FKExcel
 *
 * Copyright (c) 2006 - 2013 FKExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   FKExcel
 * @package    FKExcel_Cell
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


/** FKExcel root directory */
if (!defined('FKExcel_ROOT')) {
    /**
     * @ignore
     */
    define('FKExcel_ROOT', dirname(__FILE__) . '/../../');
    require(FKExcel_ROOT . 'FKExcel/Autoloader.php');
}


/**
 * FKExcel_Cell_DefaultValueBinder
 *
 * @category   FKExcel
 * @package    FKExcel_Cell
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_Cell_DefaultValueBinder implements FKExcel_Cell_IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  FKExcel_Cell  $cell   Cell to bind value to
     * @param  mixed          $value  Value to bind in cell
     * @return boolean
     */
    public function bindValue(FKExcel_Cell $cell, $value = null)
    {
        // sanitize UTF-8 strings
        if (is_string($value)) {
            $value = FKExcel_Shared_String::SanitizeUTF8($value);
        }

		// DD: BEGIN This block is added to original code by UNKNOWN
		// DD: Original: https://github.com/PHPOffice/FKExcel/blob/develop/Classes/FKExcel/Cell/DefaultValueBinder.php
		// DD: Fix warning in this modification crap
		// Implement your own override logic
		if (is_string($value) && isset($value[0]) && $value[0] == '0') {
			$cell->setValueExplicit($value, FKExcel_Cell_DataType::TYPE_STRING);
			return true;
		}
		// DD: END This block is added to original code by UNKNOWN

        // Set value explicit
        $cell->setValueExplicit( $value, self::dataTypeForValue($value) );

        // Done!
        return TRUE;
    }

    /**
     * DataType for value
     *
     * @param   mixed  $pValue
     * @return  string
     */
    public static function dataTypeForValue($pValue = null) {
        // Match the value against a few data types
        if (is_null($pValue)) {
            return FKExcel_Cell_DataType::TYPE_NULL;

        } elseif ($pValue === '') {
            return FKExcel_Cell_DataType::TYPE_STRING;

        } elseif ($pValue instanceof FKExcel_RichText) {
            return FKExcel_Cell_DataType::TYPE_INLINE;

        } elseif ($pValue{0} === '=' && strlen($pValue) > 1) {
            return FKExcel_Cell_DataType::TYPE_FORMULA;

        } elseif (is_bool($pValue)) {
            return FKExcel_Cell_DataType::TYPE_BOOL;

        } elseif (is_float($pValue) || is_int($pValue)) {
            return FKExcel_Cell_DataType::TYPE_NUMERIC;

        } elseif (preg_match('/^\-?([0-9]+\\.?[0-9]*|[0-9]*\\.?[0-9]+)$/', $pValue)) {
            return FKExcel_Cell_DataType::TYPE_NUMERIC;

        } elseif (is_string($pValue) && array_key_exists($pValue, FKExcel_Cell_DataType::getErrorCodes())) {
            return FKExcel_Cell_DataType::TYPE_ERROR;

        } else {
            return FKExcel_Cell_DataType::TYPE_STRING;

        }
    }
}
