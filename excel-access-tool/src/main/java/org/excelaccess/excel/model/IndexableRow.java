package org.excelaccess.excel.model;

import org.excelaccess.excel.model.annotation.ExcelInternal;
import org.excelaccess.excel.model.annotation.ExcelInternalEnum;

/**
 * Cette ligne est indexable, définissable par son numéro de ligne.
 * 
 * @author Loic Abemonty
 * 
 */
public interface IndexableRow {
    /**
     * Donne le numéro de la ligne, 0-based.
     */
    @ExcelInternal(value = ExcelInternalEnum.ROW_LINE)
    Integer getRowNum();

}
