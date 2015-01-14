package org.excelaccess.excel;

import static com.google.common.base.Preconditions.checkNotNull;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.excelaccess.excel.model.annotation.ExcelCell;
import org.excelaccess.excel.model.annotation.ExcelDocument;
import org.joda.time.DateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Classe d'accès aux lignes d'un document excel. <br/>
 * 
 * @see #parse(int, Class)
 * @author Loic Abemonty
 * 
 */
public class ExcelAccessor {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelAccessor.class);

    /**
     * Création d'une instance de ExcelAccessor d'après un fichier excel.
     * 
     * @param excelInputStream
     *            input stream du fichier Excel
     * @return une instance
     * @throws IOException
     *             si un soucis avec le flux binaire apparait
     */
    public static ExcelAccessor getInstance(InputStream excelInputStream) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(excelInputStream);
        return new ExcelAccessor(workbook);
    }

    private final ClassLoader classLoader = Thread.currentThread().getContextClassLoader();

    private final HSSFWorkbook workbook;

    /**
     * Simple constructeur avec le workbook de l'api POI.
     * 
     * @param workbook
     *            non null.
     */
    public ExcelAccessor(HSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    /**
     * Ajoute un élément dans l'excel, une ligne normalement, en fin de tableau.
     * 
     * @param <T>
     * @param clazz
     *            non null
     * @return une représentation de l'élément créé
     */
    public <T> T add(Class<T> clazz) {

        ExcelDocument excelDocumentDeclaration = clazz.getAnnotation(ExcelDocument.class);
        if (excelDocumentDeclaration == null) {
            throw new IllegalArgumentException("Class should be annotated with @ExcelDocument: " + clazz.getName());
        }
        String sheetName = excelDocumentDeclaration.sheetName();

        return add(clazz, sheetName);
    }

    /**
     * Getter du format interne de l'excel utilisé. <br/>
     * TODO loic - statuer sur le fait d'avoir ce getter.
     */
    public HSSFWorkbook getExcelWorkBook() {
        return this.workbook;
    }

    /**
     * Crée une représentation d'une ligne.
     * 
     * @param <T>
     *            la classe à instancier représentant une ligne
     * @param rowNumber
     *            numéro de la ligne, 0-based
     * @param clazz
     *            non null
     * @return une représentation de la ligne, null si la ligne n'existe pas
     */
    public <T> T parse(int rowNumber, Class<T> clazz) {
        checkNotNull(clazz, "class");

        ExcelDocument excelDocumentDeclaration = clazz.getAnnotation(ExcelDocument.class);

        if (excelDocumentDeclaration == null) {
            throw new IllegalArgumentException("Class should be annotated with @ExcelDocument: " + clazz.getName());
        }
        String sheetName = excelDocumentDeclaration.sheetName();

        return parse(clazz, sheetName, rowNumber);
    }

    /**
     * Ecrit le contenu du fichier Excel manipulé. Cette méthode de ferme pas le flux.
     * 
     * @param outputStream
     *            non null
     * @throws IOException
     *             s'il est impossible d'écrire
     */
    public void write(OutputStream outputStream) throws IOException {
        this.workbook.write(outputStream);
    }

    private <T> T parse(Class<T> clazz, String sheetName, int rowNumber) {

        int sheetIndex = this.workbook.getSheetIndex(sheetName);
        if (sheetIndex < 0) {
            sheetIndex = 0;
            LOGGER.warn("Impossible d'accéder à la feuille " + sheetName + " : utilisation de la première.");
        }
        HSSFSheet sheet = this.workbook.getSheetAt(sheetIndex);
        HSSFRow row = sheet.getRow(rowNumber);

        if (row == null) {
            // impossible de gérer une ligne qui n'existe pas.
            return null;
        }
        InvocationHandler invocationHandler = new ExcelRowInvocationHandler(clazz, row);
        T proxy = clazz.cast(Proxy.newProxyInstance(classLoader, new Class<?>[] { clazz }, invocationHandler));

        return proxy;
    }

    /**
     * Création d'une ligne et des cellules utiles de la ligne.
     * 
     * @param <T>
     * @param clazz
     *            non null
     * @param sheetName
     *            nom de la feuille
     * @return
     */
    protected <T> T add(Class<T> clazz, String sheetName) {

        int sheetIndex = this.workbook.getSheetIndex(sheetName);
        if (sheetIndex < 0) {
            sheetIndex = 0;
            LOGGER.warn("Impossible d'accéder à la feuille " + sheetName + " : utilisation de la première.");
        }
        HSSFSheet sheet = this.workbook.getSheetAt(sheetIndex);
        ExcelDocument excelDocumentDeclaration = clazz.getAnnotation(ExcelDocument.class);

        int sheetLastRow = sheet.getLastRowNum();
        if (sheetLastRow == 0 && sheet.getPhysicalNumberOfRows() == 0) {
            sheetLastRow = -1;
        }
        int sheetNewRow = excelDocumentDeclaration.startAtRow();
        if (sheetNewRow <= sheetLastRow) {
            sheetNewRow = sheetLastRow + 1;
        }
        HSSFRow createdRow = null;
        // avance un coup pour créer la prochaine ligne
        LOGGER.debug("Dernière ligne : " + sheetLastRow + " pour un démarrage à la ligne " + sheetNewRow);
        for (int i = sheetLastRow + 1; i <= sheetNewRow; i++) {
            createdRow = sheet.createRow(i);
            LOGGER.debug("Création de la ligne n°" + i + " dans la feuille " + sheetName);
        }

        if (createdRow == null) {
            LOGGER.warn("Impossible de créer une ligne dans la feuille " + sheetName);
            return null;
        }

        return parse(clazz, sheetName, createdRow.getRowNum());
        // return initializeRow(createdRow, clazz);
    }

    /**
     * Initialization d'une ligne avec création de cellule vide
     * 
     * @param <T>
     * @param row
     *            la ligne à remplir, non null
     * @param clazz
     *            , la classe qui va accéder à la ligne
     * @return l'objet représentant la ligne
     */
    protected <T> T initializeRow(HSSFRow row, Class<T> clazz) {

        // détermination de la dernière cellule
        int lastCellIdx = 0;
        for (Method method : clazz.getDeclaredMethods()) {
            ExcelCell excelCell = method.getAnnotation(ExcelCell.class);
            if (excelCell != null) {
                int cellColumnId = excelCell.value();
                if (cellColumnId > lastCellIdx) {
                    lastCellIdx = cellColumnId;
                }
                // au passage on crée les cellules qui nous intéressent avec les bons types
                Class<?> returnType = method.getReturnType();
                int cellTypeId = CELL_TYPE_BLANK;
                if (String.class.equals(returnType)) {
                    cellTypeId = CELL_TYPE_STRING;
                } else if (DateTime.class.equals(returnType)) {
                    cellTypeId = CELL_TYPE_STRING;
                } else if (int.class.equals(returnType)) {
                    cellTypeId = CELL_TYPE_NUMERIC;
                }

                row.createCell(cellColumnId, cellTypeId);
                LOGGER.debug("création de la cellule " + cellColumnId + "(" + cellTypeId + ")");
            }
        }
        // création de toutes les autres cellules au type Blank
        for (int i = 0; i < lastCellIdx; i++) {
            HSSFCell cell = row.getCell(i);
            if (cell == null) {
                row.createCell(i, CELL_TYPE_BLANK);
                LOGGER.debug("création de la cellule de remplissage " + i + "(" + CELL_TYPE_BLANK + ")");
            }
        }

        // création du proxy de la ligne
        InvocationHandler invocationHandler = new ExcelRowInvocationHandler(clazz, row);
        T proxy = clazz.cast(Proxy.newProxyInstance(classLoader, new Class<?>[] { clazz }, invocationHandler));
        return proxy;
    }
}
