package org.excelaccess.excel.model;

import java.util.Date;

import org.excelaccess.excel.model.annotation.ExcelCell;
import org.excelaccess.excel.model.annotation.ExcelCellFormat;
import org.excelaccess.excel.model.annotation.ExcelDocument;
import org.excelaccess.excel.model.annotation.RepeatableExcelCell;

/**
 * le même que {@link AxeCommandeRow} mais qui utilise des int et non pas des Integer.
 * 
 * @author Loic Abemonty
 * 
 */
@ExcelDocument(sheetName = "ProgrammeCommande-Axe", startAtRow = 3)
public interface IntAxeCommandeRow {

    @ExcelCell(3)
    String getAxe();

    @ExcelCell(4)
    String getAxeId();

    /**
     * Volume prévisionnel de la commande
     * 
     * @param index
     *            n-ième semaine de la commande, 0-based.
     */
    @RepeatableExcelCell(size = 6)
    @ExcelCell(name = "H")
    int getCommande(int index);

    @ExcelCell(1)
    String getContrat();

    @ExcelCell(name = "BX")
    @ExcelCellFormat(outputFormat = "%1$tY-%<tm-%<td", localLanguage = "fr")
    Date getDate();

    @ExcelCell(5)
    String getNature();

    void setAxe(String axe);

    void setAxeId(String axeId);

    void setCommande(int index, int volumeCommande);

    void setContrat(String contrat);

    void setDate(Date date);

    void setNature(String nature);

}
