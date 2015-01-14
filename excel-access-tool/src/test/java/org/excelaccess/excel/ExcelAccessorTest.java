package org.excelaccess.excel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.excelaccess.excel.ExcelAccessor;
import org.excelaccess.excel.model.AxeCommandeRow;
import org.excelaccess.excel.model.FluxCommandeRow;
import org.excelaccess.excel.model.IndexableAxeCommandeRow;
import org.junit.Before;
import org.junit.Test;

/**
 * Test de l'écriture dans un fichier Excel.
 * 
 * @author Loic Abemonty
 * 
 */
public class ExcelAccessorTest {

    private ExcelAccessor excelAccessor;

    @Test
    public void readAxeExcel() throws IOException {
        InputStream workbookStream = new FileInputStream("src/test/resources/excel/programme-commande.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(workbookStream);
        ExcelAccessor excelReader = new ExcelAccessor(workbook);

        AxeCommandeRow axeRow = excelReader.parse(3, AxeCommandeRow.class);
        assertEquals("1310009", axeRow.getContrat());
        assertEquals("BAYONNE>MIRAMAS", axeRow.getAxe());
    }

    @Test
    public void readCommande() {
        AxeCommandeRow axeCommandeRow = excelAccessor.parse(3, AxeCommandeRow.class);
        assertEquals("1310009", axeCommandeRow.getContrat());
        assertEquals("BAYONNE>MIRAMAS", axeCommandeRow.getAxe());
        assertEquals("87673004-87271312", axeCommandeRow.getAxeId());
        assertEquals("C", axeCommandeRow.getNature());
        assertEquals(Integer.valueOf(8), axeCommandeRow.getCommande(0));
        assertEquals(Integer.valueOf(30), axeCommandeRow.getCommande(1));
        assertEquals(Integer.valueOf(31), axeCommandeRow.getCommande(2));
        assertEquals(Integer.valueOf(32), axeCommandeRow.getCommande(3));
        assertEquals(Integer.valueOf(30), axeCommandeRow.getCommande(4));
    }

    /**
     * Test que l'on ne peut accéder à un élément répétable hors de ses bornes. <br/>
     * XXX ce NPE est assez moyen, il serait bien trouver un autre moyen.
     */
    @Test
    public void readCommandeHorsBorne() {
        AxeCommandeRow axeCommandeRow = excelAccessor.parse(3, AxeCommandeRow.class);
        // dans les limites on est bien
        assertEquals(Integer.valueOf(2), axeCommandeRow.getCommande(5));
        // un peu trop loin c'est null et ça explose en NullPointerException
        assertEquals(null, axeCommandeRow.getCommande(6));
        assertEquals(null, axeCommandeRow.getCommande(7));
    }

    /**
     * Test de la gestion de l'index négatif : réduit à zéro <br/>
     */
    @Test
    public void readCommandeHorsBorneNegative() {
        AxeCommandeRow axeCommandeRow = excelAccessor.parse(3, AxeCommandeRow.class);
        // dans les limites on est bien
        assertEquals(Integer.valueOf(2), axeCommandeRow.getCommande(5));
        // en négatif c'est géré comme un zéro (0)
        assertEquals(Integer.valueOf(8), axeCommandeRow.getCommande(-1));
    }

    /**
     * Test de lecture d'une ligne qui n'existe pas.
     */
    @Test
    public void readLineNull() {
        IndexableAxeCommandeRow indexableAxeCommandeRow = excelAccessor.parse(30, IndexableAxeCommandeRow.class);
        assertNull(indexableAxeCommandeRow);
    }

    @Before
    public void setUp() throws IOException {
        File srcFile = new File("src/test/resources/excel/programme-commande.xls");
        InputStream workbookStream = new FileInputStream(srcFile);
        HSSFWorkbook workbook = new HSSFWorkbook(workbookStream);
        excelAccessor = new ExcelAccessor(workbook);
    }

    @Test
    public void writeAxeExcel() throws IOException {
        File destFile = File.createTempFile("programme-commande", ".xls", new File("target"));
        destFile.deleteOnExit();

        AxeCommandeRow axeRow = excelAccessor.parse(3, AxeCommandeRow.class);
        assertEquals(Integer.valueOf(30), axeRow.getCommande(1));
        axeRow.setCommande(1, 10);
        axeRow.setNature("V");

        excelAccessor.write(new FileOutputStream(destFile));

        InputStream workbookStreamReread = new FileInputStream(destFile);
        HSSFWorkbook workbookReread = new HSSFWorkbook(workbookStreamReread);
        ExcelAccessor excelReaderReread = new ExcelAccessor(workbookReread);

        AxeCommandeRow axeRowReread = excelReaderReread.parse(3, AxeCommandeRow.class);
        assertEquals(Integer.valueOf(10), axeRowReread.getCommande(1));
        assertEquals("V", axeRowReread.getNature());
    }

    /**
     * Ajout une ligne dans le 2ieme onglet.
     * 
     * @throws IOException
     */
    @Test
    public void writeExcelFlux() throws IOException {
        File destFile = File.createTempFile("programme-commande", ".xls", new File("target"));
        destFile.deleteOnExit();

        FluxCommandeRow fluxRow = excelAccessor.add(FluxCommandeRow.class);
        fluxRow.setFluxId("flux-id");
        fluxRow.setFlux("MARSEILLE>ALL");
        fluxRow.setNature("C");
        fluxRow.setCommande(0, 11);

        excelAccessor.write(new FileOutputStream(destFile));

        InputStream workbookStreamReread = new FileInputStream(destFile);
        HSSFWorkbook workbookReread = new HSSFWorkbook(workbookStreamReread);
        ExcelAccessor excelReaderReread = new ExcelAccessor(workbookReread);

        FluxCommandeRow fluxRowReread = excelReaderReread.parse(4, FluxCommandeRow.class);
        assertEquals("flux-id", fluxRowReread.getFluxId());
        assertEquals("MARSEILLE>ALL", fluxRowReread.getFlux());
        assertEquals("C", fluxRowReread.getNature());
        assertEquals(Integer.valueOf(11), fluxRowReread.getCommande(0));
    }

}
