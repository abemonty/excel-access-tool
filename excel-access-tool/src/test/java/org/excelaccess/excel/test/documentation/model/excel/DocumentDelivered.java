package org.excelaccess.excel.test.documentation.model.excel;

import java.util.Date;

import org.excelaccess.excel.model.IndexableRow;
import org.excelaccess.excel.model.annotation.ExcelCell;
import org.excelaccess.excel.model.annotation.ExcelCellFormat;
import org.excelaccess.excel.model.annotation.ExcelDocument;
import org.excelaccess.excel.test.documentation.model.Document;

/**
 * Définition de l'objet Document du second onglet.
 * 
 * @author Loic Abemonty
 */
@ExcelDocument(sheetName = "Données", startAtRow = 4)
public interface DocumentDelivered extends Document, IndexableRow {

    /**
     * Nom de l'application
     */
    @Override
    @ExcelCell(value = 0)
    String getApplication();

    /**
     * Date de livraison de la version.
     */
    @Override
    @ExcelCell(name = "K")
    @ExcelCellFormat(outputExcelFormat = 0xe)
    String getDateLivraison();

    /**
     * Détails du fichier.
     */
    @Override
    @ExcelCell(name = "E")
    String getDetails();

    /**
     * Nom du fichier.
     */
    @Override
    @ExcelCell(name = "D")
    String getName();

    /**
     * Référence externe.
     */
    @Override
    @ExcelCell(name = "F")
    String getReference();

    /**
     * Catégorie de la documentation.
     */
    @Override
    @ExcelCell(value = 1)
    String getType();

    /**
     * Type du fichier (spec fonctionnelle, spec technique, etc).
     */
    @Override
    @ExcelCell(value = 2)
    String getTypeDoc();

    /**
     * Numéro de la version livrée.
     */
    @Override
    @ExcelCell(name = "J")
    Integer getVersion();

    void setApplication(String applicationName);

    void setDateLivraison(Date date);

    void setDetails(String details);

    void setName(String name);

    void setReference(String reference);

    void setType(String type);

    void setTypeDoc(String typedoc);

    void setVersion(int version);

}
