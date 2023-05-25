-- MySqlBackup.NET 2.3.8.0
-- Dump Time: 2023-05-26 01:45:43
-- --------------------------------------
-- Server version 8.0.32 MySQL Community Server - GPL


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;


-- 
-- Definition of accounts
-- 

DROP TABLE IF EXISTS `accounts`;
CREATE TABLE IF NOT EXISTS `accounts` (
  `acct_id` varchar(45) NOT NULL,
  `acct_firstname` varchar(50) DEFAULT NULL,
  `acct_lastname` varchar(50) DEFAULT NULL,
  `acct_sex` varchar(45) DEFAULT NULL,
  `acct_age` varchar(45) DEFAULT NULL,
  `acct_username` varchar(100) DEFAULT NULL,
  `acct_password` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`acct_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- 
-- Dumping data for table accounts
-- 

/*!40000 ALTER TABLE `accounts` DISABLE KEYS */;
INSERT INTO `accounts`(`acct_id`,`acct_firstname`,`acct_lastname`,`acct_sex`,`acct_age`,`acct_username`,`acct_password`) VALUES('cecile@gmail.com','Cecile','Cecile','Female','20','cecile@gmail.com','0231a1bba275eac1ebb37acb638175e1'),('doggie','doggie','aw','Male','3','doggie','79c075880be456508e77569d90bf0806'),('don@gmail.com','Don','Entong','Male','20','don@gmail.com','6a01bfa30172639e770a6aacb78a3ed4'),('donald@gmail.com','Donald','Entong','Male','21','donald@gmail.com','0d343c0f0ca763f983c8042350059f56'),('donaldmillena.entong@bicol-u.edu..ph','Donald','Don','Male','20','donaldmillena.entong@bicol-u.edu..ph','0d343c0f0ca763f983c8042350059f56'),('gemaine@gmail.com','Germaine','Fortin','Female','20','gemaine@gmail.com','b2f6a8a550241574efe30ac4ab36e7e0'),('joy@gmail.com','joy','ride','Female','16','joy@gmail.com','c2c8e798aecbc26d86e4805114b03c51'),('test','user','test','M','23','test','098f6bcd4621d373cade4e832627b4f6'),('venus@gmail,com','Venus','Rawr','Female','20','venus@gmail,com','a5717a649d346ed0c51be68888c130cd');
/*!40000 ALTER TABLE `accounts` ENABLE KEYS */;

-- 
-- Definition of customers
-- 

DROP TABLE IF EXISTS `customers`;
CREATE TABLE IF NOT EXISTS `customers` (
  `customer_id` int NOT NULL,
  `customer_name` varchar(50) DEFAULT NULL,
  `customer_sex` varchar(45) DEFAULT NULL,
  `customer_email` varchar(100) DEFAULT NULL,
  `customer_feedback` longtext,
  PRIMARY KEY (`customer_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- 
-- Dumping data for table customers
-- 

/*!40000 ALTER TABLE `customers` DISABLE KEYS */;

/*!40000 ALTER TABLE `customers` ENABLE KEYS */;

-- 
-- Definition of employees
-- 

DROP TABLE IF EXISTS `employees`;
CREATE TABLE IF NOT EXISTS `employees` (
  `employee_id` int NOT NULL,
  `employee_first_name` varchar(45) DEFAULT NULL,
  `employee_last_name` varchar(45) DEFAULT NULL,
  `employee_sex` varchar(45) DEFAULT NULL,
  `employee_age` varchar(45) DEFAULT NULL,
  `employee_email` varchar(100) DEFAULT NULL,
  `employee_address` varchar(100) DEFAULT NULL,
  `employee_postal_code` varchar(45) DEFAULT NULL,
  `employee_salary` decimal(10,2) DEFAULT NULL,
  PRIMARY KEY (`employee_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- 
-- Dumping data for table employees
-- 

/*!40000 ALTER TABLE `employees` DISABLE KEYS */;

/*!40000 ALTER TABLE `employees` ENABLE KEYS */;

-- 
-- Definition of manager
-- 

DROP TABLE IF EXISTS `manager`;
CREATE TABLE IF NOT EXISTS `manager` (
  `manager_id` int NOT NULL,
  `manager_first_name` varchar(50) DEFAULT NULL,
  `manager_last_name` varchar(50) DEFAULT NULL,
  `manager_sex` varchar(45) DEFAULT NULL,
  `manager_age` varchar(45) DEFAULT NULL,
  `manager_email` varchar(100) DEFAULT NULL,
  `manager_address` varchar(100) DEFAULT NULL,
  `manager_postal_code` varchar(45) DEFAULT NULL,
  `manager_salary` decimal(10,2) DEFAULT NULL,
  PRIMARY KEY (`manager_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- 
-- Dumping data for table manager
-- 

/*!40000 ALTER TABLE `manager` DISABLE KEYS */;

/*!40000 ALTER TABLE `manager` ENABLE KEYS */;

-- 
-- Definition of orders
-- 

DROP TABLE IF EXISTS `orders`;
CREATE TABLE IF NOT EXISTS `orders` (
  `order_id` int NOT NULL,
  `employee_id` int DEFAULT NULL,
  `customer_id` int DEFAULT NULL,
  `product_id` int DEFAULT NULL,
  `product_size` varchar(45) DEFAULT NULL,
  `product_quantity` varchar(45) DEFAULT NULL,
  `order_total_amount` decimal(10,2) DEFAULT NULL,
  `order_payment_method` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`order_id`),
  KEY `employee_id_idx` (`employee_id`),
  KEY `customer_id_idx` (`customer_id`),
  CONSTRAINT `customer_id` FOREIGN KEY (`customer_id`) REFERENCES `customers` (`customer_id`),
  CONSTRAINT `employee_id` FOREIGN KEY (`employee_id`) REFERENCES `employees` (`employee_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- 
-- Dumping data for table orders
-- 

/*!40000 ALTER TABLE `orders` DISABLE KEYS */;

/*!40000 ALTER TABLE `orders` ENABLE KEYS */;

-- 
-- Definition of products
-- 

DROP TABLE IF EXISTS `products`;
CREATE TABLE IF NOT EXISTS `products` (
  `product_id` int NOT NULL AUTO_INCREMENT,
  `product_name` varchar(100) DEFAULT NULL,
  `product_category` varchar(45) DEFAULT NULL,
  `product_price` varchar(50) DEFAULT NULL,
  `product_quantity` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`product_id`)
) ENGINE=InnoDB AUTO_INCREMENT=25 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- 
-- Dumping data for table products
-- 

/*!40000 ALTER TABLE `products` DISABLE KEYS */;
INSERT INTO `products`(`product_id`,`product_name`,`product_category`,`product_price`,`product_quantity`) VALUES(1,'Cappuccino','Coffee','60','30'),(2,'Latte','Coffee','69','30'),(3,'Espresso','Coffee','100','30'),(4,'Amerikano','Coffee','80','30'),(5,'Macchiato','Coffee','50','30'),(6,'Bagel','Bread','30','30'),(7,'Ensaimada','Bread','25','30'),(8,'Baguette','Bread','50','30'),(9,'Brioche','Bread','45','30'),(10,'Pita','Bread','40','30');
/*!40000 ALTER TABLE `products` ENABLE KEYS */;

-- 
-- Definition of sales
-- 

DROP TABLE IF EXISTS `sales`;
CREATE TABLE IF NOT EXISTS `sales` (
  `fkmanager_id` int DEFAULT NULL,
  `fkemployee_id` int DEFAULT NULL,
  `total_sales` decimal(10,2) DEFAULT NULL,
  `total_product_sales` varchar(1000) DEFAULT NULL,
  KEY `manager_id_idx` (`fkmanager_id`),
  KEY `employee_id_idx` (`fkemployee_id`),
  CONSTRAINT `fkemployee_id` FOREIGN KEY (`fkemployee_id`) REFERENCES `employees` (`employee_id`),
  CONSTRAINT `fkmanager_id` FOREIGN KEY (`fkmanager_id`) REFERENCES `manager` (`manager_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

-- 
-- Dumping data for table sales
-- 

/*!40000 ALTER TABLE `sales` DISABLE KEYS */;

/*!40000 ALTER TABLE `sales` ENABLE KEYS */;

-- 
-- Dumping functions
-- 

DROP FUNCTION IF EXISTS `avg_employee_salary`;
DELIMITER |
CREATE FUNCTION `avg_employee_salary`() RETURNS decimal(10,2)
    READS SQL DATA
    DETERMINISTIC
BEGIN
    DECLARE avg_salary DECIMAL(10,2);
    SELECT AVG(employee_salary) INTO avg_salary FROM employees;
    RETURN avg_salary;
END |
DELIMITER ;

-- 
-- Dumping procedures
-- 

DROP PROCEDURE IF EXISTS `update_product_price`;
DELIMITER |
CREATE PROCEDURE `update_product_price`(IN product_id INT, IN new_price DECIMAL(10,2))
BEGIN
    UPDATE products SET product_price = new_price WHERE product_id = product_id;
END |
DELIMITER ;

-- 
-- Dumping views
-- 

DROP TABLE IF EXISTS `customer_sales_report`;
DROP VIEW IF EXISTS `customer_sales_report`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `customer_sales_report` AS select `customers`.`customer_id` AS `customer_id`,`customers`.`customer_name` AS `customer_name`,sum(`orders`.`order_total_amount`) AS `total_sales` from (`customers` join `orders` on((`customers`.`customer_id` = `orders`.`customer_id`))) group by `customers`.`customer_id`;

DROP TABLE IF EXISTS `employee_sales_report`;
DROP VIEW IF EXISTS `employee_sales_report`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `employee_sales_report` AS select `employees`.`employee_id` AS `employee_id`,concat(`employees`.`employee_first_name`,' ',`employees`.`employee_last_name`) AS `employee_name`,sum(`orders`.`order_total_amount`) AS `total_sales` from (`employees` join `orders` on((`employees`.`employee_id` = `orders`.`employee_id`))) group by `employees`.`employee_id`;

DROP TABLE IF EXISTS `product_sales_report`;
DROP VIEW IF EXISTS `product_sales_report`;
CREATE ALGORITHM=UNDEFINED SQL SECURITY DEFINER VIEW `product_sales_report` AS select `products`.`product_category` AS `product_category`,sum(`orders`.`order_total_amount`) AS `total_sales` from (`products` join `orders` on((`products`.`product_id` = `orders`.`product_id`))) group by `products`.`product_category`;



/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;


-- Dump completed on 2023-05-26 01:45:43
-- Total time: 0:0:0:0:168 (d:h:m:s:ms)
