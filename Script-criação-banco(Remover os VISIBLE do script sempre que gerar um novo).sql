-- MySQL Script generated by MySQL Workbench
-- Thu Jun  1 16:58:26 2023
-- Model: New Model    Version: 1.0
-- MySQL Workbench Forward Engineering

SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0;
SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='ONLY_FULL_GROUP_BY,STRICT_TRANS_TABLES,NO_ZERO_IN_DATE,NO_ZERO_DATE,ERROR_FOR_DIVISION_BY_ZERO,NO_ENGINE_SUBSTITUTION';

-- -----------------------------------------------------
-- Schema rhinfo16_cf
-- -----------------------------------------------------

-- -----------------------------------------------------
-- Schema rhinfo16_cf
-- -----------------------------------------------------
CREATE SCHEMA IF NOT EXISTS `rhinfo16_cf` DEFAULT CHARACTER SET utf8 ;
USE `rhinfo16_cf` ;

-- -----------------------------------------------------
-- Table `rhinfo16_cf`.`categorias`
-- -----------------------------------------------------
CREATE TABLE IF NOT EXISTS `rhinfo16_cf`.`categorias` (
  `codigo_categoria` INT NOT NULL,
  `descricao_categoria` VARCHAR(50) NOT NULL,
  `observacao_categoria` VARCHAR(100) NOT NULL,
  `receita_ou_despesa_categoria` TINYINT(1) NOT NULL,
  PRIMARY KEY (`codigo_categoria`))
ENGINE = InnoDB;


-- -----------------------------------------------------
-- Table `rhinfo16_cf`.`receitas`
-- -----------------------------------------------------
CREATE TABLE IF NOT EXISTS `rhinfo16_cf`.`receitas` (
  `codigo_receita` INT NOT NULL,
  `descricao_receita` VARCHAR(50) NOT NULL,
  `data_pagamento` DATE NOT NULL,
  `valor_receita` DECIMAL(16,2) NOT NULL,
  `observacao_receita` VARCHAR(100) NULL,
  `pago_receita` TINYINT(1) NOT NULL,
  `categorias_codigo_categoria` INT NOT NULL,
  PRIMARY KEY (`codigo_receita`, `categorias_codigo_categoria`),
  INDEX `fk_receitas_categorias_idx` (`categorias_codigo_categoria` ASC) ,
  CONSTRAINT `fk_receitas_categorias`
    FOREIGN KEY (`categorias_codigo_categoria`)
    REFERENCES `rhinfo16_cf`.`categorias` (`codigo_categoria`)
    ON DELETE NO ACTION
    ON UPDATE NO ACTION)
ENGINE = InnoDB;


-- -----------------------------------------------------
-- Table `rhinfo16_cf`.`despesas`
-- -----------------------------------------------------
CREATE TABLE IF NOT EXISTS `rhinfo16_cf`.`despesas` (
  `codigo_despesa` INT NOT NULL,
  `descricao_despesa` VARCHAR(50) NOT NULL,
  `data_pagamento` VARCHAR(45) NOT NULL,
  `valor_despesa` DECIMAL(16,2) NOT NULL,
  `observacao_despesa` VARCHAR(100) NULL,
  `pago_despesa` TINYINT(1) NOT NULL,
  `categorias_codigo_categoria` INT NOT NULL,
  PRIMARY KEY (`codigo_despesa`, `categorias_codigo_categoria`),
  INDEX `fk_despesas_categorias1_idx` (`categorias_codigo_categoria` ASC) ,
  CONSTRAINT `fk_despesas_categorias1`
    FOREIGN KEY (`categorias_codigo_categoria`)
    REFERENCES `rhinfo16_cf`.`categorias` (`codigo_categoria`)
    ON DELETE NO ACTION
    ON UPDATE NO ACTION)
ENGINE = InnoDB;


SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS;
