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
 * @package    FKExcel_Writer
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * FKExcel_Writer_Abstract
 *
 * @category   FKExcel
 * @package    FKExcel_Writer
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
abstract class FKExcel_Writer_Abstract implements FKExcel_Writer_IWriter
{
	/**
	 * Write charts that are defined in the workbook?
	 * Identifies whether the Writer should write definitions for any charts that exist in the FKExcel object;
	 *
	 * @var	boolean
	 */
	protected $_includeCharts = FALSE;

	/**
	 * Pre-calculate formulas
	 * Forces FKExcel to recalculate all formulae in a workbook when saving, so that the pre-calculated values are
	 *    immediately available to MS Excel or other office spreadsheet viewer when opening the file
	 *
	 * @var boolean
	 */
	protected $_preCalculateFormulas = TRUE;

	/**
	 * Use disk caching where possible?
	 *
	 * @var boolean
	 */
	protected $_useDiskCaching = FALSE;

	/**
	 * Disk caching directory
	 *
	 * @var string
	 */
	protected $_diskCachingDirectory	= './';

	/**
	 * Write charts in workbook?
	 *		If this is true, then the Writer will write definitions for any charts that exist in the FKExcel object.
	 *		If false (the default) it will ignore any charts defined in the FKExcel object.
	 *
	 * @return	boolean
	 */
	public function getIncludeCharts() {
		return $this->_includeCharts;
	}

	/**
	 * Set write charts in workbook
	 *		Set to true, to advise the Writer to include any charts that exist in the FKExcel object.
	 *		Set to false (the default) to ignore charts.
	 *
	 * @param	boolean	$pValue
	 * @return	FKExcel_Writer_IWriter
	 */
	public function setIncludeCharts($pValue = FALSE) {
		$this->_includeCharts = (boolean) $pValue;
		return $this;
	}

    /**
     * Get Pre-Calculate Formulas flag
	 *     If this is true (the default), then the writer will recalculate all formulae in a workbook when saving,
	 *        so that the pre-calculated values are immediately available to MS Excel or other office spreadsheet
	 *        viewer when opening the file
	 *     If false, then formulae are not calculated on save. This is faster for saving in FKExcel, but slower
	 *        when opening the resulting file in MS Excel, because Excel has to recalculate the formulae itself
     *
     * @return boolean
     */
    public function getPreCalculateFormulas() {
    	return $this->_preCalculateFormulas;
    }

    /**
     * Set Pre-Calculate Formulas
	 *		Set to true (the default) to advise the Writer to calculate all formulae on save
	 *		Set to false to prevent precalculation of formulae on save.
     *
     * @param boolean $pValue	Pre-Calculate Formulas?
	 * @return	FKExcel_Writer_IWriter
     */
    public function setPreCalculateFormulas($pValue = TRUE) {
    	$this->_preCalculateFormulas = (boolean) $pValue;
		return $this;
    }

	/**
	 * Get use disk caching where possible?
	 *
	 * @return boolean
	 */
	public function getUseDiskCaching() {
		return $this->_useDiskCaching;
	}

	/**
	 * Set use disk caching where possible?
	 *
	 * @param 	boolean 	$pValue
	 * @param	string		$pDirectory		Disk caching directory
	 * @throws	FKExcel_Writer_Exception	when directory does not exist
	 * @return FKExcel_Writer_Excel2007
	 */
	public function setUseDiskCaching($pValue = FALSE, $pDirectory = NULL) {
		$this->_useDiskCaching = $pValue;

		if ($pDirectory !== NULL) {
    		if (is_dir($pDirectory)) {
    			$this->_diskCachingDirectory = $pDirectory;
    		} else {
    			throw new FKExcel_Writer_Exception("Directory does not exist: $pDirectory");
    		}
		}
		return $this;
	}

	/**
	 * Get disk caching directory
	 *
	 * @return string
	 */
	public function getDiskCachingDirectory() {
		return $this->_diskCachingDirectory;
	}
}
