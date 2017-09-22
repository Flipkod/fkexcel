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
 * FKExcel_CachedObjectStorage_ICache
 *
 * @category   FKExcel
 * @package    FKExcel_CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
interface FKExcel_CachedObjectStorage_ICache
{
    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param	string			$pCoord		Coordinate address of the cell to update
     * @param	FKExcel_Cell	$cell		Cell to update
	 * @return	void
     * @throws	FKExcel_Exception
     */
	public function addCacheData($pCoord, FKExcel_Cell $cell);

    /**
     * Add or Update a cell in cache
     *
     * @param	FKExcel_Cell	$cell		Cell to update
	 * @return	void
     * @throws	FKExcel_Exception
     */
	public function updateCacheData(FKExcel_Cell $cell);

    /**
     * Fetch a cell from cache identified by coordinate address
     *
     * @param	string			$pCoord		Coordinate address of the cell to retrieve
     * @return FKExcel_Cell 	Cell that was found, or null if not found
     * @throws	FKExcel_Exception
     */
	public function getCacheData($pCoord);

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @param	string			$pCoord		Coordinate address of the cell to delete
     * @throws	FKExcel_Exception
     */
	public function deleteCacheData($pCoord);

	/**
	 * Is a value set in the current FKExcel_CachedObjectStorage_ICache for an indexed cell?
	 *
	 * @param	string		$pCoord		Coordinate address of the cell to check
	 * @return	boolean
	 */
	public function isDataSet($pCoord);

	/**
	 * Get a list of all cell addresses currently held in cache
	 *
	 * @return	array of string
	 */
	public function getCellList();

	/**
	 * Get the list of all cell addresses currently held in cache sorted by column and row
	 *
	 * @return	void
	 */
	public function getSortedCellList();

	/**
	 * Clone the cell collection
	 *
	 * @param	FKExcel_Worksheet	$parent		The new worksheet
	 * @return	void
	 */
	public function copyCellCollection(FKExcel_Worksheet $parent);

	/**
	 * Identify whether the caching method is currently available
	 * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
	 *
	 * @return	boolean
	 */
	public static function cacheMethodIsAvailable();

}
