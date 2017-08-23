CREATE TABLE `tbretur` (
  `kode` CHAR(20) COLLATE utf8_general_ci NOT NULL,
  `nama` CHAR(40) COLLATE utf8_general_ci DEFAULT NULL,
  `tgl_retur` DATE DEFAULT NULL,
  `userid` CHAR(20) COLLATE utf8_general_ci DEFAULT NULL,
  `jumlah` INTEGER(11) DEFAULT NULL,
  PRIMARY KEY (`kode`) USING BTREE
) ENGINE=InnoDB
CHARACTER SET 'utf8' COLLATE 'utf8_general_ci'
;
