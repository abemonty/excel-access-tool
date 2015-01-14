package org.excelaccess.excel.test.documentation.model.excel;

import java.util.Date;

import org.excelaccess.excel.model.annotation.ExcelCell;
import org.excelaccess.excel.model.annotation.ExcelCellFormat;
import org.excelaccess.excel.model.annotation.ExcelDocument;
import org.excelaccess.excel.model.annotation.RepeatableExcelCell;

/**
 * Définition de l'objet Information de suivi du premier onglet.
 * 
 * @author Loic Abemonty
 */
@ExcelDocument(sheetName = "Intro", startAtRow = 2)
public interface InformationSuivi {

    /**
     * Intitulé de la mise à jour.
     * 
     * @param nith
     *            n-ième suivi
     * @return le texte de la mise à jour
     */
    @ExcelCell(name = "E")
    @RepeatableExcelCell(size = 5, jump = 1)
    String getSuivi(int nith);

    /**
     * Date du suivi.
     * 
     * @param nith
     *            n-ième suivi
     * @return la date du suivi
     */
    @ExcelCell(name = "F")
    @RepeatableExcelCell(size = 5, jump = 1)
    @ExcelCellFormat(localLanguage = "fr", outputExcelFormat = 0xe)
    Date getSuiviDate(int nith);

    void setSuivi(int nith, String suivi);

    void setSuiviDate(int nith, Date date);
}
