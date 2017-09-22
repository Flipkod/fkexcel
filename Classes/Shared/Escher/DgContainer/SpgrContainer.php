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
 * FKExcel_Shared_Escher_DgContainer_SpgrContainer
 *
 * @category   FKExcel
 * @package    FKExcel_Shared_Escher
 * @copyright  Copyright (c) 2006 - 2013 FKExcel (http://www.codeplex.com/FKExcel)
 */
class FKExcel_Shared_Escher_DgContainer_SpgrContainer
{
	/**
	 * Parent Shape Group Container
	 *
	 * @var FKExcel_Shared_Escher_DgContainer_SpgrContainer
	 */
	private $_parent;

	/**
	 * Shape Container collection
	 *
	 * @var array
	 */
	private $_children = array();

	/**
	 * Set parent Shape Group Container
	 *
	 * @param FKExcel_Shared_Escher_DgContainer_SpgrContainer $parent
	 */
	public function setParent($parent)
	{
		$this->_parent = $parent;
	}

	/**
	 * Get the parent Shape Group Container if any
	 *
	 * @return FKExcel_Shared_Escher_DgContainer_SpgrContainer|null
	 */
	public function getParent()
	{
		return $this->_parent;
	}

	/**
	 * Add a child. This will be either spgrContainer or spContainer
	 *
	 * @param mixed $child
	 */
	public function addChild($child)
	{
		$this->_children[] = $child;
		$child->setParent($this);
	}

	/**
	 * Get collection of Shape Containers
	 */
	public function getChildren()
	{
		return $this->_children;
	}

	/**
	 * Recursively get all spContainers within this spgrContainer
	 *
	 * @return FKExcel_Shared_Escher_DgContainer_SpgrContainer_SpContainer[]
	 */
	public function getAllSpContainers()
	{
		$allSpContainers = array();

		foreach ($this->_children as $child) {
			if ($child instanceof FKExcel_Shared_Escher_DgContainer_SpgrContainer) {
				$allSpContainers = array_merge($allSpContainers, $child->getAllSpContainers());
			} else {
				$allSpContainers[] = $child;
			}
		}

		return $allSpContainers;
	}
}
