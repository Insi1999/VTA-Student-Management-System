-- phpMyAdmin SQL Dump
-- version 4.5.1
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: Sep 29, 2022 at 12:00 PM
-- Server version: 10.1.16-MariaDB
-- PHP Version: 5.6.24

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `kasthury`
--

-- --------------------------------------------------------

--
-- Table structure for table `acc`
--

CREATE TABLE `acc` (
  `id` int(11) NOT NULL,
  `user` varchar(12) DEFAULT NULL,
  `pass` varchar(12) DEFAULT NULL,
  `nm` varchar(30) DEFAULT NULL,
  `des` varchar(15) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `acc`
--

INSERT INTO `acc` (`id`, `user`, `pass`, `nm`, `des`) VALUES
(2, 'VTA', 'root', 'Vocational Training Authority', 'Admin');

-- --------------------------------------------------------

--
-- Table structure for table `attand`
--

CREATE TABLE `attand` (
  `mis` varchar(20) NOT NULL,
  `stname` varchar(35) NOT NULL,
  `january` decimal(4,1) DEFAULT NULL,
  `cjanuary` decimal(4,1) DEFAULT NULL,
  `february` decimal(4,1) DEFAULT NULL,
  `cFebruary` decimal(4,1) DEFAULT NULL,
  `march` decimal(4,1) DEFAULT NULL,
  `cmarch` decimal(4,1) DEFAULT NULL,
  `april` decimal(4,1) DEFAULT NULL,
  `capril` decimal(4,1) DEFAULT NULL,
  `may` decimal(4,1) DEFAULT NULL,
  `cmay` decimal(4,1) DEFAULT NULL,
  `june` decimal(4,1) DEFAULT NULL,
  `cjune` decimal(4,1) DEFAULT NULL,
  `july` decimal(4,1) DEFAULT NULL,
  `cjuly` decimal(4,1) DEFAULT NULL,
  `august` decimal(4,1) DEFAULT NULL,
  `caugust` decimal(4,1) DEFAULT NULL,
  `september` decimal(4,1) DEFAULT NULL,
  `cseptember` decimal(4,1) DEFAULT NULL,
  `october` decimal(4,1) DEFAULT NULL,
  `coctober` decimal(4,1) DEFAULT NULL,
  `november` decimal(4,1) DEFAULT NULL,
  `cnovember` decimal(4,1) DEFAULT NULL,
  `december` decimal(4,1) DEFAULT NULL,
  `cdecember` decimal(4,1) DEFAULT NULL,
  `totaldays` decimal(4,1) DEFAULT NULL,
  `totalclassdays` decimal(4,1) DEFAULT NULL,
  `percentage` decimal(4,1) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `dropout`
--

CREATE TABLE `dropout` (
  `mis` varchar(20) NOT NULL,
  `stname` varchar(35) NOT NULL,
  `batch` varchar(3) DEFAULT NULL,
  `year` varchar(4) DEFAULT NULL,
  `lastatt` varchar(11) DEFAULT NULL,
  `des` varchar(15) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `final`
--

CREATE TABLE `final` (
  `mis` varchar(20) NOT NULL,
  `nic` varchar(12) NOT NULL,
  `stname` varchar(35) NOT NULL,
  `indexno` varchar(15) NOT NULL,
  `currentdetail` varchar(15) DEFAULT NULL,
  `certificate` varchar(18) NOT NULL,
  `theory` varchar(4) NOT NULL,
  `practical` varchar(4) NOT NULL,
  `nvq` varchar(8) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `modules`
--

CREATE TABLE `modules` (
  `mis` varchar(20) NOT NULL,
  `stname` varchar(35) NOT NULL,
  `m1prac` int(10) NOT NULL,
  `m1theo` int(10) NOT NULL,
  `m2prac` int(10) DEFAULT NULL,
  `m2theo` int(10) DEFAULT NULL,
  `m3prac` int(10) DEFAULT NULL,
  `m3theo` int(10) DEFAULT NULL,
  `m4prac` int(10) DEFAULT NULL,
  `m4theo` int(10) DEFAULT NULL,
  `m5prac` int(10) DEFAULT NULL,
  `m5theo` int(10) DEFAULT NULL,
  `m6prac` int(10) DEFAULT NULL,
  `m6theo` int(10) DEFAULT NULL,
  `m7prac` int(10) DEFAULT NULL,
  `m7theo` int(10) DEFAULT NULL,
  `m8prac` int(10) DEFAULT NULL,
  `m8theo` int(10) DEFAULT NULL,
  `m9prac` int(10) DEFAULT NULL,
  `m9theo` int(10) DEFAULT NULL,
  `m10prac` int(10) DEFAULT NULL,
  `m10theo` int(10) DEFAULT NULL,
  `m11prac` int(10) DEFAULT NULL,
  `m11theo` int(10) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `ojt`
--

CREATE TABLE `ojt` (
  `mis` varchar(20) NOT NULL,
  `stname` varchar(35) NOT NULL,
  `duration` varchar(8) NOT NULL,
  `ojtplace` varchar(50) NOT NULL,
  `depcontact` varchar(10) DEFAULT NULL,
  `corname` varchar(25) DEFAULT NULL,
  `corcontact` varchar(10) DEFAULT NULL,
  `salissue` varchar(3) DEFAULT NULL,
  `companytype` varchar(15) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `payments`
--

CREATE TABLE `payments` (
  `pid` int(11) NOT NULL,
  `mis` varchar(20) NOT NULL,
  `stname` varchar(35) NOT NULL,
  `ddfee` int(4) DEFAULT NULL,
  `cbtfee` int(4) DEFAULT NULL,
  `examfee` int(4) DEFAULT NULL,
  `regfee` int(4) DEFAULT NULL,
  `total` int(5) DEFAULT NULL,
  `des` varchar(15) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `stdetail`
--

CREATE TABLE `stdetail` (
  `mis` varchar(20) NOT NULL,
  `stname` varchar(35) NOT NULL,
  `nic` varchar(12) DEFAULT NULL,
  `years` varchar(4) DEFAULT NULL,
  `batch` varchar(3) NOT NULL,
  `gender` varchar(6) NOT NULL,
  `dob` date NOT NULL,
  `address` varchar(45) NOT NULL,
  `contact` varchar(10) NOT NULL,
  `science` char(1) NOT NULL,
  `maths` char(1) NOT NULL,
  `english` char(1) NOT NULL,
  `stpic` text NOT NULL,
  `doc` varchar(13) DEFAULT NULL,
  `sts` varchar(6) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Indexes for dumped tables
--

--
-- Indexes for table `acc`
--
ALTER TABLE `acc`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `uk` (`user`);

--
-- Indexes for table `attand`
--
ALTER TABLE `attand`
  ADD KEY `fk2` (`mis`);

--
-- Indexes for table `dropout`
--
ALTER TABLE `dropout`
  ADD KEY `dropfk` (`mis`);

--
-- Indexes for table `final`
--
ALTER TABLE `final`
  ADD KEY `finalfk` (`mis`);

--
-- Indexes for table `modules`
--
ALTER TABLE `modules`
  ADD KEY `mis` (`mis`);

--
-- Indexes for table `ojt`
--
ALTER TABLE `ojt`
  ADD KEY `fkojt` (`mis`);

--
-- Indexes for table `payments`
--
ALTER TABLE `payments`
  ADD PRIMARY KEY (`pid`),
  ADD KEY `mis` (`mis`);

--
-- Indexes for table `stdetail`
--
ALTER TABLE `stdetail`
  ADD PRIMARY KEY (`mis`),
  ADD UNIQUE KEY `nic` (`nic`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `acc`
--
ALTER TABLE `acc`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;
--
-- AUTO_INCREMENT for table `payments`
--
ALTER TABLE `payments`
  MODIFY `pid` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=11;
--
-- Constraints for dumped tables
--

--
-- Constraints for table `attand`
--
ALTER TABLE `attand`
  ADD CONSTRAINT `fk2` FOREIGN KEY (`mis`) REFERENCES `stdetail` (`mis`);

--
-- Constraints for table `dropout`
--
ALTER TABLE `dropout`
  ADD CONSTRAINT `dropfk` FOREIGN KEY (`mis`) REFERENCES `stdetail` (`mis`);

--
-- Constraints for table `final`
--
ALTER TABLE `final`
  ADD CONSTRAINT `finalfk` FOREIGN KEY (`mis`) REFERENCES `stdetail` (`mis`);

--
-- Constraints for table `modules`
--
ALTER TABLE `modules`
  ADD CONSTRAINT `modules_ibfk_1` FOREIGN KEY (`mis`) REFERENCES `stdetail` (`mis`);

--
-- Constraints for table `ojt`
--
ALTER TABLE `ojt`
  ADD CONSTRAINT `fkojt` FOREIGN KEY (`mis`) REFERENCES `stdetail` (`mis`);

--
-- Constraints for table `payments`
--
ALTER TABLE `payments`
  ADD CONSTRAINT `payments_ibfk_1` FOREIGN KEY (`mis`) REFERENCES `stdetail` (`mis`);

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
