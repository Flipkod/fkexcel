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
 * @package    FKExcel_Shared_Escher
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/**
 * FKExcel_Shared_Escher_DggContainer_BstoreContainer_BSE
 *
 * @category   FKExcel
 * @package    FKExcel_Shared_Escher
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_Shared_Escher_DggContainer_BstoreContainer_BSE
{
	const BLIPTYPE_ERROR	= 0x00;
	const BLIPTYPE_UNKNOWN	= 0x01;
	const BLIPTYPE_EMF		= 0x02;
	const BLIPTYPE_WMF		= 0x03;
	const BLIPTYPE_PICT		= 0x04;
	const BLIPTYPE_JPEG		= 0x05;
	const BLIPTYPE_PNG		= 0x06;
	const BLIPTYPE_DIB		= 0x07;
	const BLIPTYPE_TIFF		= 0x11;
	const BLIPTYPE_CMYKJPEG	= 0x12;

	/**
	 * The parent BLIP Store Entry Container
	 *
	 * @var FKExcel_Shared_Escher_DggContainer_BstoreContainer
	 */
	private $_parent;

	/**
	 * The BLIP (Big Large Image or Picture)
	 *
	 * @var FKExcel_Shared_Escher_DggContainer_BstoreContainer_BSE_Blip
	 */
	private $_blip;

	/**
	 * The BLIP type
	 *
	 * @var int
	 */
	private $_blipType;

	/**
	 * Set parent BLIP Store Entry Container
	 *
	 * @param FKExcel_Shared_Escher_DggContainer_BstoreContainer $parent
	 */
	public function setParent($parent)
	{
		$this->_parent = $parent;
	}

	/**
	 * Get the BLIP
	 *
	 * @return FKExcel_Shared_Escher_DggContainer_BstoreContainer_BSE_Blip
	 */
	public function getBlip()
	{
		return $this->_blip;
	}

	/**
	 * Set the BLIP
	 *
	 * @param FKExcel_Shared_Escher_DggContainer_BstoreContainer_BSE_Blip $blip
	 */
	public function setBlip($blip)
	{
		$this->_blip = $blip;
		$blip->setParent($this);
	}

	/**
	 * Get the BLIP type
	 *
	 * @return int
	 */
	public function getBlipType()
	{
		return $this->_blipType;
	}

	/**
	 * Set the BLIP type
	 *
	 * @param int
	 */
	public function setBlipType($blipType)
	{
		$this->_blipType = $blipType;
	}

}