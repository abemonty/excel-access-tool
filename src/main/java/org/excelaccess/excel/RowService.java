package org.excelaccess.excel;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.excelaccess.excel.model.IndexableRow;
import org.excelaccess.excel.model.annotation.ExcelDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Classe utilitaire pour gérer les lignes en cours d'utilisation.<br/>
 * Gestion de plusieurs types de lignes indépendemment les uns des autres.
 * 
 * @param <R>
 *            Classe représentant les données d'une ligne
 * @deprecated Utiliser RowHandler
 * @author Loic Abemonty
 */
@Deprecated
public abstract class RowService<R> {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowService.class);

    /**
     * L'accesseur à l'objet Excel.
     */
    private final ExcelAccessor excelAccessor;

    /**
     * Gestion des lignes déjà considérées, classées par type de ligne, puis par clé de ligne.<br/>
     * Il n'y a normalement qu'un seul type de ligne par RowService, mais il est possible d'en avoir plusieurs.
     */
    private final Map<Class<?>, Map<R, Integer>> typedRows = new HashMap<Class<?>, Map<R, Integer>>();

    /**
     * Constructeur.
     * 
     * @param excelAccessor
     *            l'accesseur du fichier Excel.
     */
    public RowService(ExcelAccessor excelAccessor) {
        this.excelAccessor = excelAccessor;
    }

    /**
     * Recherche la ligne correspondant à la ressource, la crée si on ne la trouve pas. <br/>
     * Si la ligne est créée, aucun champs de l'objet n'est positionné à part le numéro de ligne (
     * {@link IndexableRow#getRowNum()})
     * 
     * @param <T>
     *            la classe de la ressource (extends {@link IndexableRow}
     * @param resource
     *            la ressource
     * @param resourceKey
     *            une clé pouvant référencer la ressource
     * @return non null
     */
    public <T extends IndexableRow> T findRow(Class<T> resourceClazz, R resourceKey) {
        T row = null;
        Integer rowNumber = null;

        Map<R, Integer> rows = this.getRows(resourceClazz);
        if (rows == null) {
            rows = new HashMap<R, Integer>();
        }
        rowNumber = rows.get(resourceKey);
        if (rowNumber == null) {
            row = searchOrCreateRow(resourceClazz, resourceKey);
        } else {
            row = excelAccessor.parse(rowNumber, resourceClazz);
        }
        rows.put(resourceKey, row.getRowNum());
        this.putRows(resourceClazz, rows);

        return row;
    }

    /**
     * Lecture des lignes gérées jusqu'à présent.
     */
    public Map<R, Integer> getRows() {
        return null; // TODO faire l'union des rows
    }

    public Map<R, Integer> getRows(Class<?> resourceClazz) {
        return this.typedRows.get(resourceClazz);
    }

    public Map<R, Integer> putRows(Class<?> resourceClazz, Map<R, Integer> rows) {
        return this.typedRows.put(resourceClazz, rows);
    }

    /**
     * Méthode de vérification qu'une ressource {resource} peut être référencée par la clé {resourceKey}.<br/>
     * La ressource ici est une ligne d'un fichier excel.
     * 
     * @param <T>
     *            la classe de la ressource (extends {@link IndexableRow}
     * @param resource
     *            la ressource
     * @param resourceKey
     *            une clé pouvant représenter la ressource
     * @return
     */
    abstract protected <T extends IndexableRow> boolean compare(T resource, R resourceKey);

    protected ExcelAccessor getExcelAccessor() {
        return excelAccessor;
    }

    /**
     * Recherche d'une ligne existante avec cette ligne là, création si non trouvé.<br/>
     * La classe de la ressource doit posséder l'annotation {@link ExcelDocument}.
     * 
     * @param resourceClazz
     *            la classe de la ligne que l'on recherche
     * @param resourceKey
     *            la clé de la ligne, non null
     * @return une ligne dans l'excel, nouvelle ou ancienne.
     */
    protected <T extends IndexableRow> T searchOrCreateRow(Class<T> resourceClazz, R resourceKey) {
        T row = null;

        ExcelDocument excelDocument = resourceClazz.getAnnotation(ExcelDocument.class);
        int startRow = excelDocument.startAtRow();
        String sheetName = excelDocument.sheetName();

        HSSFWorkbook excelWorkBook = excelAccessor.getExcelWorkBook();
        int sheetIndex = excelWorkBook.getSheetIndex(sheetName);
        if (sheetIndex < 0) {
            sheetIndex = 0;
            LOGGER.warn("Impossible d'accéder à la feuille " + sheetName + " : utilisation de la première.");
        }
        HSSFSheet sheet = excelWorkBook.getSheetAt(sheetIndex);
        // "<=" car il faut inclure la dernière ligne dans la recherche
        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
            row = excelAccessor.parse(i, resourceClazz);
            if (compare(row, resourceKey)) {
                return row;
            }
        }

        row = excelAccessor.add(resourceClazz);

        return row;
    }

}