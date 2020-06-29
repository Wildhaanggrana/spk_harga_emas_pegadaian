# Host: localhost  (Version 5.5.5-10.1.13-MariaDB)
# Date: 2020-06-29 10:29:29
# Generator: MySQL-Front 5.3  (Build 5.33)

/*!40101 SET NAMES latin1 */;

#
# Structure for table "tabel_emaspegadaian"
#

DROP TABLE IF EXISTS `tabel_emaspegadaian`;
CREATE TABLE `tabel_emaspegadaian` (
  `Idemas` int(11) NOT NULL AUTO_INCREMENT,
  `emas` varchar(255) DEFAULT NULL,
  `beratemas` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`Idemas`)
) ENGINE=InnoDB AUTO_INCREMENT=1367 DEFAULT CHARSET=latin1;

#
# Data for table "tabel_emaspegadaian"
#


#
# Structure for table "tabel_peramalandanpenapsiranpegadaian"
#

DROP TABLE IF EXISTS `tabel_peramalandanpenapsiranpegadaian`;
CREATE TABLE `tabel_peramalandanpenapsiranpegadaian` (
  `idtaksiran` int(11) NOT NULL AUTO_INCREMENT,
  `type` varchar(255) DEFAULT NULL,
  `beratemas` int(255) DEFAULT NULL,
  `hargapasaran` int(11) DEFAULT NULL,
  `hargataksiran` int(11) DEFAULT NULL,
  PRIMARY KEY (`idtaksiran`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "tabel_peramalandanpenapsiranpegadaian"
#


#
# Structure for table "tabel_prediksihargaemaspegadaian"
#

DROP TABLE IF EXISTS `tabel_prediksihargaemaspegadaian`;
CREATE TABLE `tabel_prediksihargaemaspegadaian` (
  `idprediksi` varchar(255) DEFAULT NULL,
  `bulansebelumnya` varchar(255) DEFAULT NULL,
  `bulansekarang` varchar(255) DEFAULT NULL,
  `hasilprediksi` varchar(255) DEFAULT NULL,
  UNIQUE KEY `id_prediksi` (`idprediksi`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPACT;

#
# Data for table "tabel_prediksihargaemaspegadaian"
#

