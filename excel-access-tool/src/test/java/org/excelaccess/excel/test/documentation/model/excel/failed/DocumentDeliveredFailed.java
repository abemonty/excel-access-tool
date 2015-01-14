package org.excelaccess.excel.test.documentation.model.excel.failed;

import org.excelaccess.excel.model.annotation.ExcelDocument;
import org.excelaccess.excel.test.documentation.model.excel.DocumentDelivered;

/**
 * Version mal configuré de DocumentDelivered.
 * 
 * @author Loic Abemonty
 * 
 */
@ExcelDocument(sheetName = "Données", startAtRow = 122)
public interface DocumentDeliveredFailed extends DocumentDelivered {

}
