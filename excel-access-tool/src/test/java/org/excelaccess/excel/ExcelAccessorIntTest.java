package org.excelaccess.excel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.excelaccess.excel.ExcelAccessor;
import org.excelaccess.excel.model.FluxCommandeRow;
import org.excelaccess.excel.model.IntAxeCommandeRow;
import org.joda.time.DateMidnight;
import org.junit.Before;
import org.junit.Test;

/**
 * Test de l'écriture et lecture dans un fichier avec des méthodes avec int et non pas Integer.
 * 
 * @author Loic Abemonty
 * 
 */
public class ExcelAccessorIntTest {

    private ExcelAccessor excelAccessor;

    @Test
    public void lireEtEcrireDate() throws FileNotFoundException, IOException {
        File destFile = File.createTempFile("programme-commande", ".xls", new File("target"));
        destFile.deleteOnExit();

        IntAxeCommandeRow axeRow = excelAccessor.parse(3, IntAxeCommandeRow.class);
        axeRow.setDate(new DateMidnight(2011, 10, 20).toDate());

        excelAccessor.write(new FileOutputStream(destFile));

        InputStream workbookStreamReread = new FileInputStream(destFile);
        HSSFWorkbook workbookReread = new HSSFWorkbook(workbookStreamReread);
        ExcelAccessor excelReaderReread = new ExcelAccessor(workbookReread);

        IntAxeCommandeRow axeRowReread = excelReaderReread.parse(3, IntAxeCommandeRow.class);
        String dateFormatted = String.format("%1$tY-%<tm-%<td", axeRowReread.getDate());
        assertEquals("2011-10-20", dateFormatted);
    }


    @Test
    public void lireEtEcrireDateNull() throws FileNotFoundException, IOException {
        File destFile = File.createTempFile("programme-commande", ".xls", new File("target"));
        destFile.deleteOnExit();

        IntAxeCommandeRow axeRow = excelAccessor.parse(3, IntAxeCommandeRow.class);
        axeRow.setDate(new DateMidnight(2011, 10, 20).toDate());
        // la mise à null écrit bien la cellule Excel
        axeRow.setDate(null);

        excelAccessor.write(new FileOutputStream(destFile));

        InputStream workbookStreamReread = new FileInputStream(destFile);
        HSSFWorkbook workbookReread = new HSSFWorkbook(workbookStreamReread);
        ExcelAccessor excelReaderReread = new ExcelAccessor(workbookReread);

        IntAxeCommandeRow axeRowReread = excelReaderReread.parse(3, IntAxeCommandeRow.class);
        assertNull(axeRowReread.getDate());
    }

    @Test
    public void lireEtEcrireInt() throws FileNotFoundException, IOException {
        File destFile = File.createTempFile("programme-commande", ".xls", new File("target"));
        destFile.deleteOnExit();

        FluxCommandeRow fluxRow = excelAccessor.parse(3, FluxCommandeRow.class);
        fluxRow.setVersionContrat(123);

        excelAccessor.write(new FileOutputStream(destFile));

        InputStream workbookStreamReread = new FileInputStream(destFile);
        HSSFWorkbook workbookReread = new HSSFWorkbook(workbookStreamReread);
        ExcelAccessor excelReaderReread = new ExcelAccessor(workbookReread);

        FluxCommandeRow fluxRowReread = excelReaderReread.parse(3, FluxCommandeRow.class);
        Integer intFormatted = fluxRowReread.getVersionContrat();
        assertEquals(new Integer(123), intFormatted);
    }

    @Test
    public void readAxeCommande() {
        IntAxeCommandeRow axeCommandeRow = excelAccessor.parse(3, IntAxeCommandeRow.class);
        assertEquals("1310009", axeCommandeRow.getContrat());
        assertEquals("BAYONNE>MIRAMAS", axeCommandeRow.getAxe());
        assertEquals("87673004-87271312", axeCommandeRow.getAxeId());
        assertEquals("C", axeCommandeRow.getNature());
        assertEquals(8, axeCommandeRow.getCommande(0));
        assertEquals(30, axeCommandeRow.getCommande(1));
        assertEquals(31, axeCommandeRow.getCommande(2));
        assertEquals(32, axeCommandeRow.getCommande(3));
        assertEquals(30, axeCommandeRow.getCommande(4));
    }

    /**
     * Test que l'on ne peut accéder à un élément répétable hors de ses bornes. <br/>
     * XXX ce NPE est assez moyen, il serait bien trouver un autre moyen.
     */
    @Test(expected = NullPointerException.class)
    public void readCommandeHorsBorne() {
        IntAxeCommandeRow axeCommandeRow = excelAccessor.parse(3, IntAxeCommandeRow.class);
        // dans les limites on est bien
        assertEquals(2, axeCommandeRow.getCommande(5));
        // un peu trop loin c'est null et ça explose en NullPointerException
        assertEquals(null, axeCommandeRow.getCommande(6));
    }

    /**
     * Test de la gestion de l'index négatif : réduit à zéro <br/>
     */
    @Test
    public void readCommandeHorsBorneNegative() {
        IntAxeCommandeRow axeCommandeRow = excelAccessor.parse(3, IntAxeCommandeRow.class);
        // dans les limites on est bien
        assertEquals(2, axeCommandeRow.getCommande(5));
        // en négatif c'est géré comme 0 (zéro)
        assertEquals(8, axeCommandeRow.getCommande(-1));
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

        IntAxeCommandeRow axeRow = excelAccessor.parse(3, IntAxeCommandeRow.class);
        assertEquals(30, axeRow.getCommande(1));
        axeRow.setCommande(1, 10);
        axeRow.setNature("V");

        excelAccessor.write(new FileOutputStream(destFile));

        InputStream workbookStreamReread = new FileInputStream(destFile);
        HSSFWorkbook workbookReread = new HSSFWorkbook(workbookStreamReread);
        ExcelAccessor excelReaderReread = new ExcelAccessor(workbookReread);

        IntAxeCommandeRow axeRowReread = excelReaderReread.parse(3, IntAxeCommandeRow.class);
        assertEquals(10, axeRowReread.getCommande(1));
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

        IntAxeCommandeRow axeRow = excelAccessor.add(IntAxeCommandeRow.class);
        axeRow.setAxeId("flux-id");
        axeRow.setAxe("MARSEILLE>ALL");
        axeRow.setNature("C");
        axeRow.setCommande(0, 11);

        excelAccessor.write(new FileOutputStream(destFile));

        InputStream workbookStreamReread = new FileInputStream(destFile);
        HSSFWorkbook workbookReread = new HSSFWorkbook(workbookStreamReread);
        ExcelAccessor excelReaderReread = new ExcelAccessor(workbookReread);

        IntAxeCommandeRow fluxRowReread = excelReaderReread.parse(4, IntAxeCommandeRow.class);
        assertEquals("flux-id", fluxRowReread.getAxeId());
        assertEquals("MARSEILLE>ALL", fluxRowReread.getAxe());
        assertEquals("C", fluxRowReread.getNature());
        assertEquals(11, fluxRowReread.getCommande(0));
    }
}
