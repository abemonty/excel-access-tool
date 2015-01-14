package org.excelaccess.excel;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
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
 * @param <K>
 *            Classe identifiant une ligne
 * @param <L>
 *            Classe représentant les données d'une ligne
 * 
 * @author Loic Abemonty
 */
public abstract class RowHandler<K, L extends IndexableRow> {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowHandler.class);

    /**
     * L'accesseur à l'objet Excel.
     */
    private final ExcelAccessor excelAccessor;

    /**
     * Gestion des lignes déjà considérées, classées par clé de ligne.
     */
    private final Map<K, Integer> rows = new HashMap<K, Integer>();

    private Class<L> resourceClass;

    private Class<K> resourceKeyClass;

    /**
     * Constructeur.
     * 
     * @param excelAccessor
     *            l'accesseur du fichier Excel.
     */
    public RowHandler(ExcelAccessor excelAccessor) {
        this.excelAccessor = excelAccessor;
        findParameterizedClass();
    }

    public Map<K, Integer> addRows(Map<K, Integer> rows) {
        this.rows.putAll(rows);
        return this.rows;
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
    public L findRow(K resourceKey) {
        L row = null;
        Integer rowNumber = null;

        Map<K, Integer> rows = this.getRows();

        rowNumber = rows.get(resourceKey);
        if (rowNumber == null) {
            row = searchOrCreateRow(resourceKey);
        } else {
            row = excelAccessor.parse(rowNumber, this.resourceClass);
        }
        rows.put(resourceKey, row.getRowNum());
        this.addRows(rows);

        return row;
    }

    /**
     * Fournit la première ligne de la zone définie par {@link #resourceClass}.
     * 
     * @return null si elle n'existe pas.
     */
    public L firstRow() {
        L row = null;

        ExcelDocument excelDocument = this.resourceClass.getAnnotation(ExcelDocument.class);
        if (excelDocument != null) {
            int startAtRow = excelDocument.startAtRow();
            return this.excelAccessor.parse(startAtRow, resourceClass);
        }

        return row;
    }

    @SuppressWarnings("unchecked")
    private void findParameterizedClass() {
        if (this.resourceClass == null) {
            Type type = ((ParameterizedType) getClass().getGenericSuperclass()).getActualTypeArguments()[1];
            // manage the parameterized entity type.
            if (type instanceof ParameterizedType) {
                this.resourceClass = (Class<L>) ((ParameterizedType) type).getRawType();
            } else {
                this.resourceClass = (Class<L>) type;
            }
        }
        if (this.resourceKeyClass == null) {
            Type type = ((ParameterizedType) getClass().getGenericSuperclass()).getActualTypeArguments()[0];
            // manage the parameterized entity type.
            if (type instanceof ParameterizedType) {
                this.resourceKeyClass = (Class<K>) ((ParameterizedType) type).getRawType();
            } else {
                this.resourceKeyClass = (Class<K>) type;
            }
        }
    }

    protected ExcelAccessor getExcelAccessor() {
        return excelAccessor;
    }

    /**
     * Lecture des lignes gérées jusqu'à présent.
     */
    protected Map<K, Integer> getRows() {
        return this.rows;
    }

    /**
     * Méthode de vérification qu'une ressource {resource} peut être référencée par la clé {resourceKey}.<br/>
     * La ressource ici est une ligne d'un fichier excel.
     * 
     * @param resource
     *            la ressource
     * @param resourceKey
     *            une clé pouvant représenter la ressource
     * @return true si la clé correspond à la ressource
     */
    abstract protected boolean identifyResource(L resource, K resourceKey);

    /**
     * Compare 2 objets, en passant par la méthode compare de l'attribut1, si l'attribut1 est null alors la méthode est
     * passante et retourne true.
     * 
     * @param <M>
     *            le type, quelconque, de l'attribut1
     * @param <N>
     *            le type, quelconque, de l'attribut
     * @param attribut1
     *            peut être null
     * @param attribut2
     *            peut être null
     * @return true si ils sont égaux, false sinon
     */
    protected <M, N> boolean isMatching(M attribut1, N attribut2) {
        if ((attribut1 == null && attribut2 == null) || (attribut1 == null && attribut2 != null)) {
            return true;
        }
        // if (attribut1 != null && attribut2 == null) {
        // return false;
        // }

        return attribut1.equals(attribut2);
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
    protected L searchOrCreateRow(K resourceKey) {
        L row = null;

        ExcelDocument excelDocument = this.resourceClass.getAnnotation(ExcelDocument.class);
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
            row = excelAccessor.parse(i, this.resourceClass);
            if (identifyResource(row, resourceKey)) {
                return row;
            }
        }

        row = excelAccessor.add(this.resourceClass);

        return row;
    }
}