package org.excelaccess.excel.model;

import org.excelaccess.excel.model.annotation.ExcelCell;
import org.excelaccess.excel.model.annotation.ExcelDocument;
import org.excelaccess.excel.model.annotation.RepeatableExcelCell;

@ExcelDocument(sheetName = "ProgrammeCommande-Flux", startAtRow = 3)
public interface FluxCommandeRow extends CommandeRow {

    /**
     * Volume prévisionnel de la commande
     * 
     * @param index
     *            n-ième semaine de la commande, 0-based.
     */

    @ExcelCell(name = "A")
    String getClient();

    @Override
    @RepeatableExcelCell(size = 6)
    @ExcelCell(name = "J")
    Integer getCommande(int index);

    @ExcelCell(name = "F")
    String getFlux();

    @ExcelCell(name = "G")
    String getFluxId();

    @ExcelCell(name = "H")
    String getNature();

    @ExcelCell(name = "B")
    String getNumContrat();

    @ExcelCell(name = "C")
    Integer getVersionContrat();

    void setClient(String client);

    @Override
    void setCommande(int index, Integer volumeCommande);

    void setFlux(String fluxLabel);

    void setFluxId(String fluxId);

    void setNature(String nature);

    void setNumContrat(String numContrat);

    void setVersionContrat(Integer versionContrat);

}
