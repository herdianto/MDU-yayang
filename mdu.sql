-- phpMyAdmin SQL Dump
-- version 4.5.1
-- http://www.phpmyadmin.net
--
-- Host: 127.0.0.1
-- Generation Time: Sep 05, 2016 at 01:54 PM
-- Server version: 10.1.13-MariaDB
-- PHP Version: 5.5.37

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `mdu`
--

-- --------------------------------------------------------

--
-- Table structure for table `material`
--

CREATE TABLE `material` (
  `Code` varchar(10) NOT NULL,
  `Name` varchar(30) NOT NULL,
  `Unit` varchar(30) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `material`
--

INSERT INTO `material` (`Code`, `Name`, `Unit`) VALUES
('00005', 'trafo', 'buah'),
('00006', 'kabel twist', 'meter'),
('007', 'Hmmmm', 'hmmmm');

-- --------------------------------------------------------

--
-- Table structure for table `transaction`
--

CREATE TABLE `transaction` (
  `ID` int(10) NOT NULL,
  `userID` varchar(10) NOT NULL,
  `Code` varchar(10) NOT NULL,
  `Date` date NOT NULL,
  `Qty` int(5) NOT NULL,
  `Tug10No` varchar(30) DEFAULT NULL,
  `Tug9No` varchar(30) DEFAULT NULL,
  `Condition` enum('1','2','3') NOT NULL,
  `GoodIssueNo` varchar(30) DEFAULT NULL,
  `PKLGNo` varchar(30) DEFAULT NULL,
  `Remark` varchar(30) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `transaction`
--

INSERT INTO `transaction` (`ID`, `userID`, `Code`, `Date`, `Qty`, `Tug10No`, `Tug9No`, `Condition`, `GoodIssueNo`, `PKLGNo`, `Remark`) VALUES
(1, '', '00006', '2016-09-14', 2323, '123123', NULL, '2', NULL, 'asda', 'asdasd'),
(2, '', '00006', '2016-09-14', 2323, '123123', NULL, '2', NULL, 'asda', 'asdasd'),
(3, 'yayang', '00006', '2016-09-14', 10, 'asasad', NULL, '2', NULL, 'asdasd', 'asdasdasdasdas'),
(4, 'yayang', '00005', '2016-09-14', 10, 'asasad', NULL, '2', NULL, 'asdasd', 'asdasdasdasdas'),
(5, 'yayang', '007', '2016-09-14', 44, 'fsgsgs', NULL, '2', NULL, 'ffdgdfg', 'dgdfgdffdg'),
(6, '', '007', '2016-09-14', 123, 'qweqe', NULL, '1', '1221', NULL, NULL),
(7, '', '00006', '2016-09-14', 12, 'asd', NULL, '1', 'adasasad', NULL, NULL),
(8, '', '00006', '2016-09-14', 22, 'asdad', NULL, '1', '2312', NULL, NULL),
(9, '', '007', '2016-09-14', 1, 'qweqw', NULL, '1', 'qwwe', NULL, NULL),
(10, '', '00006', '2016-09-14', 221, 'asdasd', NULL, '1', 'lalala', NULL, NULL),
(11, '', '00006', '2016-09-14', 221, 'asdasd', NULL, '3', NULL, '', ''),
(12, '', '00006', '2016-09-14', 221, 'asdasd', NULL, '3', NULL, '', ''),
(13, '', '00006', '2016-09-14', 134, '12123', NULL, '2', NULL, 'werwer', 'werwerwerwer'),
(14, 'yayang', '00006', '2016-09-14', 12, 'lalalal', NULL, '2', NULL, 'hahah', 'ahahahha'),
(15, 'yayang', '007', '2016-09-14', 12, 'test', NULL, '3', NULL, 'aaa', 'aaaaaaaaaa'),
(16, 'yayang', '00006', '2016-09-14', 45, '2aasd', NULL, '2', NULL, 'PK vsdfsdf', '4234234234sdsdsdf');

-- --------------------------------------------------------

--
-- Table structure for table `user`
--

CREATE TABLE `user` (
  `UserID` varchar(10) NOT NULL,
  `Password` varchar(32) NOT NULL,
  `LastLogin` datetime DEFAULT NULL,
  `AccessRight` enum('0','1','2') NOT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `user`
--

INSERT INTO `user` (`UserID`, `Password`, `LastLogin`, `AccessRight`) VALUES
('yayang', 'yayang', NULL, '1'),
('anggi', 'anggi', NULL, '0');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `material`
--
ALTER TABLE `material`
  ADD PRIMARY KEY (`Code`);

--
-- Indexes for table `transaction`
--
ALTER TABLE `transaction`
  ADD PRIMARY KEY (`ID`),
  ADD KEY `userID` (`userID`),
  ADD KEY `Code` (`Code`);

--
-- Indexes for table `user`
--
ALTER TABLE `user`
  ADD PRIMARY KEY (`UserID`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `transaction`
--
ALTER TABLE `transaction`
  MODIFY `ID` int(10) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=17;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
