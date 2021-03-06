<?php
// @codingStandardsIgnoreFile
// @codeCoverageIgnoreStart
// this is an autogenerated file - do not edit
spl_autoload_register(
	function($class) {
		static $classes = null;
		if ($classes === null) {
			$classes = array(
				'xsp\\templateexcel\\templateexcel' => '/templateExcel.class.php',
				'phpexcel_iofactory' => '/../lib/PHPExcel/Classes/PHPExcel/IOFactory.php'
				
				);
		}
		$cn = strtolower($class);
		// var_dump($cn );
		if (isset($classes[$cn])) {
			require __DIR__ . $classes[$cn];
			
		}
	},
	true,
	false
	);