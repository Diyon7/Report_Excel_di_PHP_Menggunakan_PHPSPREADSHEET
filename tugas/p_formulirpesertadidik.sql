-- phpMyAdmin SQL Dump
-- version 4.9.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Apr 21, 2020 at 04:40 PM
-- Server version: 10.4.8-MariaDB
-- PHP Version: 7.3.10

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `p_formulirpesertadidik`
--

-- --------------------------------------------------------

--
-- Table structure for table `formulir`
--

CREATE TABLE `formulir` (
  `id` int(11) NOT NULL,
  `input` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE current_timestamp(),
  `nama` varchar(50) NOT NULL,
  `jk` varchar(15) NOT NULL,
  `nisn` varchar(25) NOT NULL,
  `nik` varchar(50) NOT NULL,
  `tempatlahir` varchar(50) NOT NULL,
  `tanggallahir` date NOT NULL,
  `nomorakta` varchar(50) NOT NULL,
  `agama` varchar(50) NOT NULL,
  `negara` varchar(50) NOT NULL,
  `kebutuhankhusus` varchar(15) NOT NULL,
  `alamat` varchar(80) NOT NULL,
  `rt` varchar(5) NOT NULL,
  `rw` varchar(5) NOT NULL,
  `dusun` varchar(25) NOT NULL,
  `desa` varchar(25) NOT NULL,
  `kecamatan` varchar(25) NOT NULL,
  `kp` varchar(15) NOT NULL,
  `lintang` varchar(25) NOT NULL,
  `bujur` varchar(25) NOT NULL,
  `temting` varchar(15) NOT NULL,
  `modtrans` varchar(25) NOT NULL,
  `kks` varchar(50) DEFAULT NULL,
  `anakke` varchar(5) NOT NULL,
  `pekps` varchar(15) NOT NULL,
  `kps` varchar(50) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Dumping data for table `formulir`
--

INSERT INTO `formulir` (`id`, `input`, `nama`, `jk`, `nisn`, `nik`, `tempatlahir`, `tanggallahir`, `nomorakta`, `agama`, `negara`, `kebutuhankhusus`, `alamat`, `rt`, `rw`, `dusun`, `desa`, `kecamatan`, `kp`, `lintang`, `bujur`, `temting`, `modtrans`, `kks`, `anakke`, `pekps`, `kps`) VALUES
(37, '2020-04-14 11:55:19', 'Muhamad Frendi Diyon Ardiansah', 'Laki-laki', '978675857586868687879', '967687676868675768699', 'Mojokerto', '1999-06-09', '9767875758686869', 'islam', 'Indonesia', 'Tidak-Kebutuhan', 'Jl. Raya Ngambar', '18', '05', 'Ngambar', 'Bambe', 'Driyorejo', '61177', '7A° sampai 8A° Lintang Se', '112A° sampai 113A° Bujur ', 'orang_tua', 'kendaraan_pribadi', '', '02', 'Tidak', ''),
(75, '2020-04-21 14:11:41', 'Muhamad Frendi Diyon Ardiansah', 'Laki-laki', '978675857586868687879', '6868997899790087979', 'Mojokerto', '1999-06-09', '787676588897787', 'islam', 'Indonesia', 'Tidak-Kebutuhan', 'Jl. Raya Bambe', '18', '05', 'Ngambar', 'Bambe', 'Driyorejo', '61177', '7A° sampai 8A° Lintang Se', '112A° sampai 113A° Bujur', 'orang_tua', 'kendaraan_pribadi', '', '02', 'Tidak', '');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `formulir`
--
ALTER TABLE `formulir`
  ADD PRIMARY KEY (`id`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `formulir`
--
ALTER TABLE `formulir`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=76;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
