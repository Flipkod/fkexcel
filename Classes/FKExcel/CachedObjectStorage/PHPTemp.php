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
 * @package    FKExcel_CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * FKExcel_CachedObjectStorage_PHPTemp
 *
 * @category   FKExcel
 * @package    FKExcel_CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_CachedObjectStorage_PHPTemp extends FKExcel_CachedObjectStorage_CacheBase implements FKExcel_CachedObjectStorage_ICache {

	/**
	 * Name of the file for this cache
	 *
	 * @var string
	 */
	private $_fileHandle = null;

	/**
	 * Memory limit to use before reverting to file cache
	 *
	 * @var integer
	 */
	private $_memoryCacheSize = null;

    /**
     * Store cell data in cache for the current cell object if it's "dirty",
     *     and the 'nullify' the current cell object
     *
	 * @return	void
     * @throws	FKExcel_Exception
     */
	protected function _storeData() {
		if ($this->_currentCellIsDirty) {
			$this->_currentObject->detach();

			fseek($this->_fileHandle,0,SEEK_END);
			$offset = ftell($this->_fileHandle);
			fwrite($this->_fileHandle, serialize($this->_currentObject));
			$this->_cellCache[$this->_currentObjectID]	= array('ptr' => $offset,
																'sz'  => ftell($this->_fileHandle) - $offset
															   );
			$this->_currentCellIsDirty = false;
		}
		$this->_currentObjectID = $this->_currentObject = null;
	}	//	function _storeData()


    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param	string			$pCoord		Coordinate address of the cell to update
     * @param	FKExcel_Cell	$cell		Cell to update
	 * @return	void
     * @throws	FKExcel_Exception
     */
	public function addCacheData($pCoord, FKExcel_Cell $cell) {
		if (($pCoord !== $this->_currentObjectID) && ($this->_currentObjectID !== null)) {
			$this->_storeData();
		}

		$this->_currentObjectID = $pCoord;
		$this->_currentObject = $cell;
		$this->_currentCellIsDirty = true;

		return $cell;
	}	//	function addCacheData()


    /**
     * Get cell at a specific coordinate
     *
     * @param 	string 			$pCoord		Coordinate of the cell
     * @throws 	FKExcel_Exception
     * @return 	FKExcel_Cell 	Cell that was found, or null if not found
     */
	public function getCacheData($pCoord) {
		if ($pCoord === $this->_currentObjectID) {
			return $this->_currentObject;
		}
		$this->_storeData();

		//	Check if the entry that has been requested actually exists
		if (!isset($this->_cellCache[$pCoord])) {
			//	Return null if requested entry doesn't exist in cache
			return null;
		}

		//	Set current entry to the requested entry
		$this->_currentObjectID = $pCoord;
		fseek($this->_fileHandle,$this->_cellCache[$pCoord]['ptr']);
		$this->_currentObject = unserialize(fread($this->_fileHandle,$this->_cellCache[$pCoord]['sz']));
        //    Re-attach this as the cell's parent
        $this->_currentObject->attach($this);

		//	Return requested entry
		return $this->_currentObject;
	}	//	function getCacheData()


	/**
	 * Get a list of all cell addresses currently held in cache
	 *
	 * @return  array of string
	 */
	public function getCellList() {
		if ($this->_currentObjectID !== null) {
			$this->_storeData();
		}

		return parent::getCellList();
	}


	/**
	 * Clone the cell collection
	 *
	 * @param	FKExcel_Worksheet	$parent		The new worksheet
	 * @return	void
	 */
	public function copyCellCollection(FKExcel_Worksheet $parent) {
		parent::copyCellCollection($parent);
		//	Open a new stream for the cell cache data
		$newFileHandle = fopen('php://temp/maxmemory:'.$this->_memoryCacheSize,'a+');
		//	Copy the existing cell cache data to the new stream
		fseek($this->_fileHandle,0);
		while (!feof($this->_fileHandle)) {
			fwrite($newFileHandle,fread($this->_fileHandle, 1024));
		}
		$this->_fileHandle = $newFileHandle;
	}	//	function copyCellCollection()


	/**
	 * Clear the cell collection and disconnect from our parent
	 *
	 * @return	void
	 */
	public function unsetWorksheetCells() {
		if(!is_null($this->_currentObject)) {
			$this->_currentObject->detach();
			$this->_currentObject = $this->_currentObjectID = null;
		}
		$this->_cellCache = array();

		//	detach ourself from the worksheet, so that it can then delete this object successfully
		$this->_parent = null;

		//	Close down the php://temp file
		$this->__destruct();
	}	//	function unsetWorksheetCells()


	/**
	 * Initialise this new cell collection
	 *
	 * @param	FKExcel_Worksheet	$parent		The worksheet for this cell collection
	 * @param	array of mixed		$arguments	Additional initialisation arguments
	 */
	public function __construct(FKExcel_Worksheet $parent, $arguments) {
		$this->_memoryCacheSize	= (isset($arguments['memoryCacheSize']))	? $arguments['memoryCacheSize']	: '1MB';

		parent::__construct($parent);
		if (is_null($this->_fileHandle)) {
			$this->_fileHandle = fopen('php://temp/maxmemory:'.$this->_memoryCacheSize,'a+');
		}
	}	//	function __construct()


	/**
	 * Destroy this cell collection
	 */
	public function __destruct() {
		if (!is_null($this->_fileHandle)) {
			fclose($this->_fileHandle);
		}
		$this->_fileHandle = null;
	}	//	function __destruct()

}
