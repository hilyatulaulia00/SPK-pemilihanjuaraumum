# Host: localhost  (Version 5.5.5-10.4.11-MariaDB)
# Date: 2020-06-29 20:14:43
# Generator: MySQL-Front 5.3  (Build 5.33)

/*!40101 SET NAMES utf8 */;

#
# Structure for table "login"
#

DROP TABLE IF EXISTS `login`;
CREATE TABLE `login` (
  `nmadmin` text NOT NULL,
  `Password` int(11) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "login"
#


#
# Structure for table "tb_guru"
#

DROP TABLE IF EXISTS `tb_guru`;
CREATE TABLE `tb_guru` (
  `nip` int(11) NOT NULL AUTO_INCREMENT,
  `nmguru` varchar(255) DEFAULT NULL,
  `jabatan` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`nip`)
) ENGINE=InnoDB AUTO_INCREMENT=1798009 DEFAULT CHARSET=latin1;

#
# Data for table "tb_guru"
#

INSERT INTO `tb_guru` VALUES (1703055,'Kasma','Honorer'),(1703098,'Sukma ','PNS'),(1798008,'Masnawati','PNS');

#
# Structure for table "tb_kriteria"
#

DROP TABLE IF EXISTS `tb_kriteria`;
CREATE TABLE `tb_kriteria` (
  `kdkriteria` char(5) NOT NULL DEFAULT '',
  `nmkriteria` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`kdkriteria`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "tb_kriteria"
#

INSERT INTO `tb_kriteria` VALUES ('C1','Rata Rata Nilai'),('C2','Prestasi Bidang Olahraga'),('C3','Prestasi Bidang Seni dan Cerdas Cermat'),('C4','Keikutsertaan kegiatan ekskul'),('C5','Etika disekolah');

#
# Structure for table "tb_penilaian"
#

DROP TABLE IF EXISTS `tb_penilaian`;
CREATE TABLE `tb_penilaian` (
  `nis` int(11) NOT NULL,
  `nmsiswa` varchar(30) DEFAULT NULL,
  `nip` int(11) DEFAULT NULL,
  `nmguru` varchar(30) DEFAULT NULL,
  `unsurpenilaian` varchar(30) DEFAULT NULL,
  `kdkriteria` int(11) NOT NULL,
  `nmkriteria` varchar(15) DEFAULT NULL,
  `nilai` int(11) DEFAULT NULL,
  PRIMARY KEY (`kdkriteria`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "tb_penilaian"
#


#
# Structure for table "tb_siswa"
#

DROP TABLE IF EXISTS `tb_siswa`;
CREATE TABLE `tb_siswa` (
  `nis` int(11) NOT NULL AUTO_INCREMENT,
  `nmsiswa` varchar(30) DEFAULT NULL,
  `kelas` char(8) DEFAULT NULL,
  `jnskel` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`nis`)
) ENGINE=InnoDB AUTO_INCREMENT=14124 DEFAULT CHARSET=latin1;

#
# Data for table "tb_siswa"
#

INSERT INTO `tb_siswa` VALUES (1231,'Musdalifah','VIII.3','perempuan'),(1232,'Andi Dermawan','VIII.3','laki-laki'),(1233,'Annisa Aqila','VIII.3','perempuan'),(1236,'Navya Azzahrah','VIII.3','perempuan');
