-- phpMyAdmin SQL Dump
-- version 5.1.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Mar 22, 2022 at 11:07 AM
-- Server version: 10.4.21-MariaDB
-- PHP Version: 8.0.12

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `baan_sort`
--

-- --------------------------------------------------------

--
-- Table structure for table `sorted_table`
--

CREATE TABLE `sorted_table` (
  `sorted_id` int(11) NOT NULL,
  `groupNo` int(11) NOT NULL,
  `baan` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- --------------------------------------------------------

--
-- Table structure for table `student`
--

CREATE TABLE `student` (
  `student_id` int(11) NOT NULL,
  `groupNo` int(11) NOT NULL,
  `fullName` varchar(255) NOT NULL,
  `nickName` varchar(255) NOT NULL,
  `sex` varchar(255) NOT NULL,
  `line_id` varchar(255) NOT NULL,
  `phone` varchar(255) NOT NULL,
  `bloodType` varchar(255) NOT NULL,
  `foodAllergies` varchar(255) NOT NULL,
  `dietRestrictions` varchar(255) NOT NULL,
  `drugAllergies` varchar(255) NOT NULL,
  `otherAllergies` varchar(255) NOT NULL,
  `congenitalDiseases` varchar(255) NOT NULL,
  `emergency` varchar(255) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

--
-- Dumping data for table `student`
--


--
-- Indexes for dumped tables
--

--
-- Indexes for table `sorted_table`
--
ALTER TABLE `sorted_table`
  ADD PRIMARY KEY (`sorted_id`);

--
-- Indexes for table `student`
--
ALTER TABLE `student`
  ADD PRIMARY KEY (`student_id`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `sorted_table`
--
ALTER TABLE `sorted_table`
  MODIFY `sorted_id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `student`
--
ALTER TABLE `student`
  MODIFY `student_id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=1;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
