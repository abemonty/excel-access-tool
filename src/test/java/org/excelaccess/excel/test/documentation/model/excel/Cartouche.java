package org.excelaccess.excel.test.documentation.model.excel;

import java.math.BigDecimal;
import java.util.Date;
import org.excelaccess.excel.model.annotation.ExcelCell;
import org.excelaccess.excel.model.annotation.ExcelCellFormat;
import org.excelaccess.excel.model.annotation.ExcelDocument;

/**
 * Définition de l'objet Cartouche du premier onglet.
 * 
 * @author Loic Abemonty
 */
@ExcelDocument(sheetName = "Intro", startAtRow = 2)
public interface Cartouche {
  
  /**
   * Intitulé de l'action.
   * 
   * @return le texte de l'action
   */
  @ExcelCell(value = 2)
  String getAction();
  
  /**
   * La date de l'action.
   * 
   * @return une date
   */
  @ExcelCell(name = "B")
  @ExcelCellFormat(outputExcelFormat = 0xe)
  Date getDate();
  
  /**
   * Le coût de l'action.
   * 
   * @return un texte
   */
  @ExcelCell(name = "D")
  BigDecimal getCost();
  
  /**
   * Le nom de l'utilisateur.
   * 
   * @return son nom
   */
  @ExcelCell(name = "A")
  String getUserName();
  
  void setAction(String action);
  
  void setDate(Date date);
  
  void setCost(BigDecimal cost);
  
  void setUserName(String userName);
}
