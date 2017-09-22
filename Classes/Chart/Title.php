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
 * @category	FKExcel
 * @package		FKExcel_Chart
 * @copyright	Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 * @license		http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version		##VERSION##, ##DATE##
 */


/**
 * FKExcel_Chart_Title
 *
 * @category	FKExcel
 * @package		FKExcel_Chart
 * @copyright	Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_Chart_Title
{

	/**
	 * Title Caption
	 *
	 * @var string
	 */
	private $_caption = null;

	/**
	 * Title Layout
	 *
	 * @var FKExcel_Chart_Layout
	 */
	private $_layout = null;

	/**
	 * Create a new FKExcel_Chart_Title
	 */
	public function __construct($caption = null, FKExcel_Chart_Layout $layout = null)
	{
		$this->_caption = $caption;
		$this->_layout = $layout;
	}

	/**
	 * Get caption
	 *
	 * @return string
	 */
	public function getCaption() {
		return $this->_caption;
	}

	/**
	 * Set caption
	 *
	 * @param string $caption
	 */
	public function setCaption($caption = null) {
		$this->_caption = $caption;
	}

	/**
	 * Get Layout
	 *
	 * @return FKExcel_Chart_Layout
	 */
	public function getLayout() {
		return $this->_layout;
	}

}
