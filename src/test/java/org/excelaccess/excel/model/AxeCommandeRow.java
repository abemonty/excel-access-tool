package org.excelaccess.excel.model;

import org.excelaccess.excel.model.annotation.ExcelCell;
import org.excelaccess.excel.model.annotation.ExcelDocument;
import org.excelaccess.excel.model.annotation.RepeatableExcelCell;

@ExcelDocument(sheetName = "ProgrammeCommande-Axe", startAtRow = 3)
public interface AxeCommandeRow extends CommandeRow {

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
    @Override
    @RepeatableExcelCell(size = 6)
    @ExcelCell(name = "H")
    Integer getCommande(int index);

    @ExcelCell(1)
    String getContrat();

    @ExcelCell(5)
    String getNature();

    void setAxe(String axe);

    void setAxeId(String axeId);

    @Override
    void setCommande(int index, Integer volumeCommande);

    void setContrat(String contrat);

    void setNature(String nature);
}
