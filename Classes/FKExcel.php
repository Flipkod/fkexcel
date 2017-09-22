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
 * @package    FKExcel
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


/** FKExcel root directory */
if (!defined('FKExcel_ROOT')) {
    define('FKExcel_ROOT', dirname(__FILE__) . '/');
    require(FKExcel_ROOT . 'Classes/Autoloader.php');
}

/**
 * FKExcel
 *
 * @category   FKExcel
 * @package    FKExcel
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel
{
    /**
     * Unique ID
     *
     * @var string
     */
    private $_uniqueID;

    /**
     * Document properties
     *
     * @var FKExcel_DocumentProperties
     */
    private $_properties;

    /**
     * Document security
     *
     * @var FKExcel_DocumentSecurity
     */
    private $_security;

    /**
     * Collection of Worksheet objects
     *
     * @var FKExcel_Worksheet[]
     */
    private $_workSheetCollection = array();

    /**
     * Calculation Engine
     *
     * @var FKExcel_Calculation
     */
    private $_calculationEngine = NULL;

    /**
     * Active sheet index
     *
     * @var int
     */
    private $_activeSheetIndex = 0;

    /**
     * Named ranges
     *
     * @var FKExcel_NamedRange[]
     */
    private $_namedRanges = array();

    /**
     * CellXf supervisor
     *
     * @var FKExcel_Style
     */
    private $_cellXfSupervisor;

    /**
     * CellXf collection
     *
     * @var FKExcel_Style[]
     */
    private $_cellXfCollection = array();

    /**
     * CellStyleXf collection
     *
     * @var FKExcel_Style[]
     */
    private $_cellStyleXfCollection = array();

    /**
     * Create a new FKExcel with one Worksheet
     */
    public function __construct()
    {
        $this->_uniqueID = uniqid();
        $this->_calculationEngine	= FKExcel_Calculation::getInstance($this);

        // Initialise worksheet collection and add one worksheet
        $this->_workSheetCollection = array();
        $this->_workSheetCollection[] = new FKExcel_Worksheet($this);
        $this->_activeSheetIndex = 0;

        // Create document properties
        $this->_properties = new FKExcel_DocumentProperties();

        // Create document security
        $this->_security = new FKExcel_DocumentSecurity();

        // Set named ranges
        $this->_namedRanges = array();

        // Create the cellXf supervisor
        $this->_cellXfSupervisor = new FKExcel_Style(true);
        $this->_cellXfSupervisor->bindParent($this);

        // Create the default style
        $this->addCellXf(new FKExcel_Style);
        $this->addCellStyleXf(new FKExcel_Style);
    }

    /**
     * Code to execute when this worksheet is unset()
     *
     */
    public function __destruct() {
        FKExcel_Calculation::unsetInstance($this);
        $this->disconnectWorksheets();
    }    //    function __destruct()

    /**
     * Disconnect all worksheets from this FKExcel workbook object,
     *    typically so that the FKExcel object can be unset
     *
     */
    public function disconnectWorksheets()
    {
        $worksheet = NULL;
        foreach($this->_workSheetCollection as $k => &$worksheet) {
            $worksheet->disconnectCells();
            $this->_workSheetCollection[$k] = null;
        }
        unset($worksheet);
        $this->_workSheetCollection = array();
    }

    /**
     * Return the calculation engine for this worksheet
     *
     * @return FKExcel_Calculation
     */
    public function getCalculationEngine()
    {
        return $this->_calculationEngine;
    }	//	function getCellCacheController()

    /**
     * Get properties
     *
     * @return FKExcel_DocumentProperties
     */
    public function getProperties()
    {
        return $this->_properties;
    }

    /**
     * Set properties
     *
     * @param FKExcel_DocumentProperties    $pValue
     */
    public function setProperties(FKExcel_DocumentProperties $pValue)
    {
        $this->_properties = $pValue;
    }

    /**
     * Get security
     *
     * @return FKExcel_DocumentSecurity
     */
    public function getSecurity()
    {
        return $this->_security;
    }

    /**
     * Set security
     *
     * @param FKExcel_DocumentSecurity    $pValue
     */
    public function setSecurity(FKExcel_DocumentSecurity $pValue)
    {
        $this->_security = $pValue;
    }

    /**
     * Get active sheet
     *
     * @return FKExcel_Worksheet
     */
    public function getActiveSheet()
    {
        return $this->_workSheetCollection[$this->_activeSheetIndex];
    }

    /**
     * Create sheet and add it to this workbook
     *
     * @param  int|null $iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @return FKExcel_Worksheet
     * @throws FKExcel_Exception
     */
    public function createSheet($iSheetIndex = NULL)
    {
        $newSheet = new FKExcel_Worksheet($this);
        $this->addSheet($newSheet, $iSheetIndex);
        return $newSheet;
    }

    /**
     * Check if a sheet with a specified name already exists
     *
     * @param  string $pSheetName  Name of the worksheet to check
     * @return boolean
     */
    public function sheetNameExists($pSheetName)
    {
        return ($this->getSheetByName($pSheetName) !== NULL);
    }

    /**
     * Add sheet
     *
     * @param  FKExcel_Worksheet $pSheet
     * @param  int|null $iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @return FKExcel_Worksheet
     * @throws FKExcel_Exception
     */
    public function addSheet(FKExcel_Worksheet $pSheet, $iSheetIndex = NULL)
    {
        if ($this->sheetNameExists($pSheet->getTitle())) {
            throw new FKExcel_Exception(
                "Workbook already contains a worksheet named '{$pSheet->getTitle()}'. Rename this worksheet first."
            );
        }

        if($iSheetIndex === NULL) {
            if ($this->_activeSheetIndex < 0) {
                $this->_activeSheetIndex = 0;
            }
            $this->_workSheetCollection[] = $pSheet;
        } else {
            // Insert the sheet at the requested index
            array_splice(
                $this->_workSheetCollection,
                $iSheetIndex,
                0,
                array($pSheet)
            );

            // Adjust active sheet index if necessary
            if ($this->_activeSheetIndex >= $iSheetIndex) {
                ++$this->_activeSheetIndex;
            }
        }
        return $pSheet;
    }

    /**
     * Remove sheet by index
     *
     * @param  int $pIndex Active sheet index
     * @throws FKExcel_Exception
     */
    public function removeSheetByIndex($pIndex = 0)
    {

        $numSheets = count($this->_workSheetCollection);

        if ($pIndex > $numSheets - 1) {
            throw new FKExcel_Exception(
                "You tried to remove a sheet by the out of bounds index: {$pIndex}. The actual number of sheets is {$numSheets}."
            );
        } else {
            array_splice($this->_workSheetCollection, $pIndex, 1);
        }
        // Adjust active sheet index if necessary
        if (($this->_activeSheetIndex >= $pIndex) &&
            ($pIndex > count($this->_workSheetCollection) - 1)) {
            --$this->_activeSheetIndex;
        }

    }

    /**
     * Get sheet by index
     *
     * @param  int $pIndex Sheet index
     * @return FKExcel_Worksheet
     * @throws FKExcel_Exception
     */
    public function getSheet($pIndex = 0)
    {

        $numSheets = count($this->_workSheetCollection);

        if ($pIndex > $numSheets - 1) {
            throw new FKExcel_Exception(
                "Your requested sheet index: {$pIndex} is out of bounds. The actual number of sheets is {$numSheets}."
            );
        } else {
            return $this->_workSheetCollection[$pIndex];
        }
    }

    /**
     * Get all sheets
     *
     * @return FKExcel_Worksheet[]
     */
    public function getAllSheets()
    {
        return $this->_workSheetCollection;
    }

    /**
     * Get sheet by name
     *
     * @param  string $pName Sheet name
     * @return FKExcel_Worksheet
     */
    public function getSheetByName($pName = '')
    {
        $worksheetCount = count($this->_workSheetCollection);
        for ($i = 0; $i < $worksheetCount; ++$i) {
            if ($this->_workSheetCollection[$i]->getTitle() === $pName) {
                return $this->_workSheetCollection[$i];
            }
        }

        return NULL;
    }

    /**
     * Get index for sheet
     *
     * @param  FKExcel_Worksheet $pSheet
     * @return Sheet index
     * @throws FKExcel_Exception
     */
    public function getIndex(FKExcel_Worksheet $pSheet)
    {
        foreach ($this->_workSheetCollection as $key => $value) {
            if ($value->getHashCode() == $pSheet->getHashCode()) {
                return $key;
            }
        }

        throw new FKExcel_Exception("Sheet does not exist.");
    }

    /**
     * Set index for sheet by sheet name.
     *
     * @param  string $sheetName Sheet name to modify index for
     * @param  int $newIndex New index for the sheet
     * @return New sheet index
     * @throws FKExcel_Exception
     */
    public function setIndexByName($sheetName, $newIndex)
    {
        $oldIndex = $this->getIndex($this->getSheetByName($sheetName));
        $pSheet = array_splice(
            $this->_workSheetCollection,
            $oldIndex,
            1
        );
        array_splice(
            $this->_workSheetCollection,
            $newIndex,
            0,
            $pSheet
        );
        return $newIndex;
    }

    /**
     * Get sheet count
     *
     * @return int
     */
    public function getSheetCount()
    {
        return count($this->_workSheetCollection);
    }

    /**
     * Get active sheet index
     *
     * @return int Active sheet index
     */
    public function getActiveSheetIndex()
    {
        return $this->_activeSheetIndex;
    }

    /**
     * Set active sheet index
     *
     * @param  int $pIndex Active sheet index
     * @throws FKExcel_Exception
     * @return FKExcel_Worksheet
     */
    public function setActiveSheetIndex($pIndex = 0)
    {
        $numSheets = count($this->_workSheetCollection);

        if ($pIndex > $numSheets - 1) {
            throw new FKExcel_Exception(
                "You tried to set a sheet active by the out of bounds index: {$pIndex}. The actual number of sheets is {$numSheets}."
            );
        } else {
            $this->_activeSheetIndex = $pIndex;
        }
        return $this->getActiveSheet();
    }

    /**
     * Set active sheet index by name
     *
     * @param  string $pValue Sheet title
     * @return FKExcel_Worksheet
     * @throws FKExcel_Exception
     */
    public function setActiveSheetIndexByName($pValue = '')
    {
        if (($worksheet = $this->getSheetByName($pValue)) instanceof FKExcel_Worksheet) {
            $this->setActiveSheetIndex($this->getIndex($worksheet));
            return $worksheet;
        }

        throw new FKExcel_Exception('Workbook does not contain sheet:' . $pValue);
    }

    /**
     * Get sheet names
     *
     * @return string[]
     */
    public function getSheetNames()
    {
        $returnValue = array();
        $worksheetCount = $this->getSheetCount();
        for ($i = 0; $i < $worksheetCount; ++$i) {
            $returnValue[] = $this->getSheet($i)->getTitle();
        }

        return $returnValue;
    }

    /**
     * Add external sheet
     *
     * @param  FKExcel_Worksheet $pSheet External sheet to add
     * @param  int|null $iSheetIndex Index where sheet should go (0,1,..., or null for last)
     * @throws FKExcel_Exception
     * @return FKExcel_Worksheet
     */
    public function addExternalSheet(FKExcel_Worksheet $pSheet, $iSheetIndex = null) {
        if ($this->sheetNameExists($pSheet->getTitle())) {
            throw new FKExcel_Exception("Workbook already contains a worksheet named '{$pSheet->getTitle()}'. Rename the external sheet first.");
        }

        // count how many cellXfs there are in this workbook currently, we will need this below
        $countCellXfs = count($this->_cellXfCollection);

        // copy all the shared cellXfs from the external workbook and append them to the current
        foreach ($pSheet->getParent()->getCellXfCollection() as $cellXf) {
            $this->addCellXf(clone $cellXf);
        }

        // move sheet to this workbook
        $pSheet->rebindParent($this);

        // update the cellXfs
        foreach ($pSheet->getCellCollection(false) as $cellID) {
            $cell = $pSheet->getCell($cellID);
            $cell->setXfIndex( $cell->getXfIndex() + $countCellXfs );
        }

        return $this->addSheet($pSheet, $iSheetIndex);
    }

    /**
     * Get named ranges
     *
     * @return FKExcel_NamedRange[]
     */
    public function getNamedRanges() {
        return $this->_namedRanges;
    }

    /**
     * Add named range
     *
     * @param  FKExcel_NamedRange $namedRange
     * @return FKExcel
     */
    public function addNamedRange(FKExcel_NamedRange $namedRange) {
        if ($namedRange->getScope() == null) {
            // global scope
            $this->_namedRanges[$namedRange->getName()] = $namedRange;
        } else {
            // local scope
            $this->_namedRanges[$namedRange->getScope()->getTitle().'!'.$namedRange->getName()] = $namedRange;
        }
        return true;
    }

    /**
     * Get named range
     *
     * @param  string $namedRange
     * @param  FKExcel_Worksheet|null $pSheet Scope. Use null for global scope
     * @return FKExcel_NamedRange|null
     */
    public function getNamedRange($namedRange, FKExcel_Worksheet $pSheet = null) {
        $returnValue = null;

        if ($namedRange != '' && ($namedRange !== NULL)) {
            // first look for global defined name
            if (isset($this->_namedRanges[$namedRange])) {
                $returnValue = $this->_namedRanges[$namedRange];
            }

            // then look for local defined name (has priority over global defined name if both names exist)
            if (($pSheet !== NULL) && isset($this->_namedRanges[$pSheet->getTitle() . '!' . $namedRange])) {
                $returnValue = $this->_namedRanges[$pSheet->getTitle() . '!' . $namedRange];
            }
        }

        return $returnValue;
    }

    /**
     * Remove named range
     *
     * @param  string  $namedRange
     * @param  FKExcel_Worksheet|null  $pSheet  Scope: use null for global scope.
     * @return FKExcel
     */
    public function removeNamedRange($namedRange, FKExcel_Worksheet $pSheet = null) {
        if ($pSheet === NULL) {
            if (isset($this->_namedRanges[$namedRange])) {
                unset($this->_namedRanges[$namedRange]);
            }
        } else {
            if (isset($this->_namedRanges[$pSheet->getTitle() . '!' . $namedRange])) {
                unset($this->_namedRanges[$pSheet->getTitle() . '!' . $namedRange]);
            }
        }
        return $this;
    }

    /**
     * Get worksheet iterator
     *
     * @return FKExcel_WorksheetIterator
     */
    public function getWorksheetIterator() {
        return new FKExcel_WorksheetIterator($this);
    }

    /**
     * Copy workbook (!= clone!)
     *
     * @return FKExcel
     */
    public function copy() {
        $copied = clone $this;

        $worksheetCount = count($this->_workSheetCollection);
        for ($i = 0; $i < $worksheetCount; ++$i) {
            $this->_workSheetCollection[$i] = $this->_workSheetCollection[$i]->copy();
            $this->_workSheetCollection[$i]->rebindParent($this);
        }

        return $copied;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        foreach($this as $key => $val) {
            if (is_object($val) || (is_array($val))) {
                $this->{$key} = unserialize(serialize($val));
            }
        }
    }

    /**
     * Get the workbook collection of cellXfs
     *
     * @return FKExcel_Style[]
     */
    public function getCellXfCollection()
    {
        return $this->_cellXfCollection;
    }

    /**
     * Get cellXf by index
     *
     * @param  int $pIndex
     * @return FKExcel_Style
     */
    public function getCellXfByIndex($pIndex = 0)
    {
        return $this->_cellXfCollection[$pIndex];
    }

    /**
     * Get cellXf by hash code
     *
     * @param  string $pValue
     * @return FKExcel_Style|false
     */
    public function getCellXfByHashCode($pValue = '')
    {
        foreach ($this->_cellXfCollection as $cellXf) {
            if ($cellXf->getHashCode() == $pValue) {
                return $cellXf;
            }
        }
        return false;
    }

    /**
     * Check if style exists in style collection
     *
     * @param  FKExcel_Style $pCellStyle
     * @return boolean
     */
    public function cellXfExists($pCellStyle = null)
    {
        return in_array($pCellStyle, $this->_cellXfCollection, true);
    }

    /**
     * Get default style
     *
     * @return FKExcel_Style
     * @throws FKExcel_Exception
     */
    public function getDefaultStyle()
    {
        if (isset($this->_cellXfCollection[0])) {
            return $this->_cellXfCollection[0];
        }
        throw new FKExcel_Exception('No default style found for this workbook');
    }

    /**
     * Add a cellXf to the workbook
     *
     * @param FKExcel_Style $style
     */
    public function addCellXf(FKExcel_Style $style)
    {
        $this->_cellXfCollection[] = $style;
        $style->setIndex(count($this->_cellXfCollection) - 1);
    }

    /**
     * Remove cellXf by index. It is ensured that all cells get their xf index updated.
     *
     * @param  int $pIndex Index to cellXf
     * @throws FKExcel_Exception
     */
    public function removeCellXfByIndex($pIndex = 0)
    {
        if ($pIndex > count($this->_cellXfCollection) - 1) {
            throw new FKExcel_Exception("CellXf index is out of bounds.");
        } else {
            // first remove the cellXf
            array_splice($this->_cellXfCollection, $pIndex, 1);

            // then update cellXf indexes for cells
            foreach ($this->_workSheetCollection as $worksheet) {
                foreach ($worksheet->getCellCollection(false) as $cellID) {
                    $cell = $worksheet->getCell($cellID);
                    $xfIndex = $cell->getXfIndex();
                    if ($xfIndex > $pIndex ) {
                        // decrease xf index by 1
                        $cell->setXfIndex($xfIndex - 1);
                    } else if ($xfIndex == $pIndex) {
                        // set to default xf index 0
                        $cell->setXfIndex(0);
                    }
                }
            }
        }
    }

    /**
     * Get the cellXf supervisor
     *
     * @return FKExcel_Style
     */
    public function getCellXfSupervisor()
    {
        return $this->_cellXfSupervisor;
    }

    /**
     * Get the workbook collection of cellStyleXfs
     *
     * @return FKExcel_Style[]
     */
    public function getCellStyleXfCollection()
    {
        return $this->_cellStyleXfCollection;
    }

    /**
     * Get cellStyleXf by index
     *
     * @param  int $pIndex
     * @return FKExcel_Style
     */
    public function getCellStyleXfByIndex($pIndex = 0)
    {
        return $this->_cellStyleXfCollection[$pIndex];
    }

    /**
     * Get cellStyleXf by hash code
     *
     * @param  string $pValue
     * @return FKExcel_Style|false
     */
    public function getCellStyleXfByHashCode($pValue = '')
    {
        foreach ($this->_cellXfStyleCollection as $cellStyleXf) {
            if ($cellStyleXf->getHashCode() == $pValue) {
                return $cellStyleXf;
            }
        }
        return false;
    }

    /**
     * Add a cellStyleXf to the workbook
     *
     * @param FKExcel_Style $pStyle
     */
    public function addCellStyleXf(FKExcel_Style $pStyle)
    {
        $this->_cellStyleXfCollection[] = $pStyle;
        $pStyle->setIndex(count($this->_cellStyleXfCollection) - 1);
    }

    /**
     * Remove cellStyleXf by index
     *
     * @param int $pIndex
     * @throws FKExcel_Exception
     */
    public function removeCellStyleXfByIndex($pIndex = 0)
    {
        if ($pIndex > count($this->_cellStyleXfCollection) - 1) {
            throw new FKExcel_Exception("CellStyleXf index is out of bounds.");
        } else {
            array_splice($this->_cellStyleXfCollection, $pIndex, 1);
        }
    }

    /**
     * Eliminate all unneeded cellXf and afterwards update the xfIndex for all cells
     * and columns in the workbook
     */
    public function garbageCollect()
    {
        // how many references are there to each cellXf ?
        $countReferencesCellXf = array();
        foreach ($this->_cellXfCollection as $index => $cellXf) {
            $countReferencesCellXf[$index] = 0;
        }

        foreach ($this->getWorksheetIterator() as $sheet) {

            // from cells
            foreach ($sheet->getCellCollection(false) as $cellID) {
                $cell = $sheet->getCell($cellID);
                ++$countReferencesCellXf[$cell->getXfIndex()];
            }

            // from row dimensions
            foreach ($sheet->getRowDimensions() as $rowDimension) {
                if ($rowDimension->getXfIndex() !== null) {
                    ++$countReferencesCellXf[$rowDimension->getXfIndex()];
                }
            }

            // from column dimensions
            foreach ($sheet->getColumnDimensions() as $columnDimension) {
                ++$countReferencesCellXf[$columnDimension->getXfIndex()];
            }
        }

        // remove cellXfs without references and create mapping so we can update xfIndex
        // for all cells and columns
        $countNeededCellXfs = 0;
        foreach ($this->_cellXfCollection as $index => $cellXf) {
            if ($countReferencesCellXf[$index] > 0 || $index == 0) { // we must never remove the first cellXf
                ++$countNeededCellXfs;
            } else {
                unset($this->_cellXfCollection[$index]);
            }
            $map[$index] = $countNeededCellXfs - 1;
        }
        $this->_cellXfCollection = array_values($this->_cellXfCollection);

        // update the index for all cellXfs
        foreach ($this->_cellXfCollection as $i => $cellXf) {
            $cellXf->setIndex($i);
        }

        // make sure there is always at least one cellXf (there should be)
        if (empty($this->_cellXfCollection)) {
            $this->_cellXfCollection[] = new FKExcel_Style();
        }

        // update the xfIndex for all cells, row dimensions, column dimensions
        foreach ($this->getWorksheetIterator() as $sheet) {

            // for all cells
            foreach ($sheet->getCellCollection(false) as $cellID) {
                $cell = $sheet->getCell($cellID);
                $cell->setXfIndex( $map[$cell->getXfIndex()] );
            }

            // for all row dimensions
            foreach ($sheet->getRowDimensions() as $rowDimension) {
                if ($rowDimension->getXfIndex() !== null) {
                    $rowDimension->setXfIndex( $map[$rowDimension->getXfIndex()] );
                }
            }

            // for all column dimensions
            foreach ($sheet->getColumnDimensions() as $columnDimension) {
                $columnDimension->setXfIndex( $map[$columnDimension->getXfIndex()] );
            }

            // also do garbage collection for all the sheets
            $sheet->garbageCollect();
        }
    }

    /**
     * Return the unique ID value assigned to this spreadsheet workbook
     *
     * @return string
     */
    public function getID() {
        return $this->_uniqueID;
    }

}