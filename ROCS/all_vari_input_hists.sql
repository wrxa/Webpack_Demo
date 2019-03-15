-- phpMyAdmin SQL Dump
-- version 3.2.4
-- http://www.phpmyadmin.net
--
-- ホスト: localhost
-- 生成時間: 2018 年 3 月 26 日 10:27
-- サーバのバージョン: 5.1.41
-- PHP のバージョン: 5.3.1

SET SQL_MODE="NO_AUTO_VALUE_ON_ZERO";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

--
-- データベース: `rocs`
--

-- --------------------------------------------------------

--
-- テーブルの構造 `all_vari_input_hists`
--

CREATE TABLE IF NOT EXISTS `all_vari_input_hists` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `CAR_ID` int(11) NOT NULL,
  `FACTOR_ID` int(3) NOT NULL,
  `KEY1` char(100) COLLATE utf8_unicode_ci NOT NULL,
  `KEY2` char(100) COLLATE utf8_unicode_ci NOT NULL,
  `KEY3` char(100) COLLATE utf8_unicode_ci NOT NULL,
  `KEY4` char(100) COLLATE utf8_unicode_ci NOT NULL,
  `KEY5` char(100) COLLATE utf8_unicode_ci NOT NULL,
  `VALUE1` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE2` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE3` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE4` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE5` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE6` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE7` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE8` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE9` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE10` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE11` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE12` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE13` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE14` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE15` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE16` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE17` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE18` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE19` varchar(256) COLLATE utf8_unicode_ci,
  `VALUE20` varchar(256) COLLATE utf8_unicode_ci,
  `UPDATE_USER_ID` char(8) COLLATE utf8_unicode_ci NOT NULL,
  `UPDATE_DATE` timestamp,
  PRIMARY KEY (`ID`,`CAR_ID`,`FACTOR_ID`,`KEY1`,`KEY2`,`KEY3`,`KEY4`,`KEY5`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci AUTO_INCREMENT=1;
--
-- テーブルのデータをダンプしています `all_vari_input_hists`
--


/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
