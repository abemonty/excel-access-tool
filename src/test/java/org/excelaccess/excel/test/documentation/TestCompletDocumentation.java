package org.excelaccess.excel.test.documentation;

import static org.junit.Assert.assertEquals;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.excelaccess.excel.ExcelAccessor;
import org.excelaccess.excel.test.documentation.model.excel.Cartouche;
import org.excelaccess.excel.test.documentation.model.excel.DocumentDelivered;
import org.excelaccess.excel.test.documentation.model.excel.InformationSuivi;
import org.excelaccess.excel.test.documentation.model.excel.failed.DocumentDeliveredFailed;
import org.excelaccess.excel.test.documentation.service.DocumentKey;
import org.excelaccess.excel.test.documentation.service.DocumentRowService;
import org.excelaccess.excel.test.documentation.service.failed.DocumentFailedRowService;
import org.joda.time.DateMidnight;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Test que l'on veut le plus complet possible.<br/>
 * Lecture d'un fichier excel sur 2 onglets, lecture de date, de champs qui se répètent, ajout de lignes.
 * 
 * @author Loic Abemonty
 */
public class TestCompletDocumentation {
  
  private ExcelAccessor excelAccessor;
  
  private static final Logger LOGGER = LoggerFactory.getLogger(TestCompletDocumentation.class);
  
  /**
   * Initialisation de l'accesseur.
   */
  @Before
  public void avantTout() throws IOException {
    File srcFile = new File("src/test/resources/excel/test_cartouche_data.xls");
    InputStream workbookStream = new FileInputStream(srcFile);
    HSSFWorkbook workbook = new HSSFWorkbook(workbookStream);
    excelAccessor = new ExcelAccessor(workbook);
  }
  
  /**
   * Lecture des valeurs de la première ligne.
   */
  @Test
  public void lectureDonnees_Document_firstRow() {
    
    DocumentRowService documentRowService = new DocumentRowService(excelAccessor);
    DocumentDelivered documentDelivered = documentRowService.firstRow();
    
    assertEquals(Integer.valueOf(4), documentDelivered.getRowNum());
    assertEquals("MLMC", documentDelivered.getApplication());
    assertEquals("PIL", documentDelivered.getType());
    assertEquals("DOCI", documentDelivered.getTypeDoc());
    assertEquals("Plan de gestion de la configuration", documentDelivered.getName());
    assertEquals(null, documentDelivered.getDetails());
    assertEquals(null, documentDelivered.getReference());
    assertEquals(null, documentDelivered.getVersion());
  }
  
  /**
   * Lecture des valeurs de la première ligne qui n'existe pas.
   */
  @Test
  public void lectureDonnees_Document_firstRow_KO() {
    
    DocumentFailedRowService documentRowService = new DocumentFailedRowService(excelAccessor);
    DocumentDeliveredFailed documentDeliveredFailed = documentRowService.firstRow();
    
    assertEquals(null, documentDeliveredFailed);
  }
  
  /**
   * Recherche d'une ligne avec des contraintes mal utilisées.
   */
  @Test
  public void lectureDonnees_Document_rechercheChampsVideMalUtilise() {
    DocumentKey documentKey = new DocumentKey();
    documentKey.application = "MLMC";
    documentKey.type = "PIL";
    documentKey.typeDoc = "DOCi";
    documentKey.name = "Manuel interne de livraison d’une version (génération, packaging, livraison)";
    // mauvaise utilisation d'un champ de filtre à vide
    documentKey.details = "";
    
    DocumentRowService documentRowService = new DocumentRowService(excelAccessor);
    DocumentDelivered documentDelivered = documentRowService.findRow(documentKey);
    
    assertEquals(Integer.valueOf(82), documentDelivered.getRowNum());
  }
  
  /**
   * Recherche d'une ligne avec peu de contrainte.
   */
  @Test
  public void lectureDonnees_Document_rechercheSimple() {
    DocumentKey documentKey = new DocumentKey();
    documentKey.application = "MLMC";
    documentKey.type = "PIL";
    documentKey.typeDoc = "DOCi";
    documentKey.name = "Manuel interne de livraison d’une version (génération, packaging, livraison)";
    
    DocumentRowService documentRowService = new DocumentRowService(excelAccessor);
    DocumentDelivered documentDelivered = documentRowService.findRow(documentKey);
    
    assertEquals(Integer.valueOf(5), documentDelivered.getRowNum());
  }
  
  /**
   * Lecture des valeurs d'une ligne.
   */
  @Test
  public void lectureDonnees_Document_simple() {
    
    DocumentDelivered documentDelivered = excelAccessor.parse(5, DocumentDelivered.class);
    
    assertEquals(Integer.valueOf(5), documentDelivered.getRowNum());
    assertEquals("MLMC", documentDelivered.getApplication());
    assertEquals("PIL", documentDelivered.getType());
    assertEquals("DOCi", documentDelivered.getTypeDoc());
    assertEquals("Manuel interne de livraison d’une version (génération, packaging, livraison)", documentDelivered.getName());
    assertEquals(null, documentDelivered.getDetails());
    assertEquals(null, documentDelivered.getReference());
    assertEquals(null, documentDelivered.getVersion());
  }
  
  /**
   * Accès en lecture/écriture des informations de suivi qui se répètent.
   * 
   * @throws IOException
   */
  @Test
  public void lectureEcriture_informationSuivi() throws IOException {
    
    int rowNumber = 2;
    DateMidnight dateMinuit = new DateMidnight(2012, 1, 1);
    // lecture des informations
    for (int i = 0; i < 4; i++) {
      InformationSuivi informationSuivi = excelAccessor.parse(rowNumber + i, InformationSuivi.class);
      assertEquals("info maj " + (i + 1) + "." + (i + 1), informationSuivi.getSuivi(i));
      assertEquals(dateMinuit.plusDays(i).toDate(), informationSuivi.getSuiviDate(i));
    }
    // complétion de la 5ieme ligne
    InformationSuivi informationSuivi5th = excelAccessor.parse(rowNumber - 1 + 5, InformationSuivi.class);
    for (int i = 0; i < 5; i++) {
      informationSuivi5th.setSuiviDate(i, dateMinuit.toDate());
    }
    informationSuivi5th.setSuivi(4, "info maj 5.5");
    
    // sauvegarde du résultat
    File outputFile = File.createTempFile("test_cartouche_data", ".xls", new File("target"));
    FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
    excelAccessor.write(fileOutputStream);
    fileOutputStream.close();
    
    // relecture de l'ensemble
    ExcelAccessor newExcelFileAccessor = ExcelAccessor.getInstance(new FileInputStream(outputFile));
    InformationSuivi informationSuivi = newExcelFileAccessor.parse(rowNumber - 1 + 5, InformationSuivi.class);
    // relecture des informations
    for (int i = 0; i < 4; i++) {
      assertEquals("info maj 5." + (i + 1), informationSuivi.getSuivi(i));
      assertEquals(dateMinuit.toDate(), informationSuivi.getSuiviDate(i));
    }
  }
  
  /**
   * Lecture, écriture et relecture des valeurs d'une ligne.
   * 
   * @throws IOException
   */
  @Test
  public void lectureEcritureDonnees_Document_simple() throws IOException {
    
    DocumentDelivered documentDelivered = excelAccessor.parse(5, DocumentDelivered.class);
    
    documentDelivered.setDetails("");
    documentDelivered.setReference("NPI_DCV_145");
    documentDelivered.setVersion(13);
    
    // sauvegarde du résultat
    File outputFile = File.createTempFile("test_cartouche_data", ".xls", new File("target"));
    FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
    excelAccessor.write(fileOutputStream);
    fileOutputStream.close();
    
    // relecture de l'ensemble
    ExcelAccessor newExcelFileAccessor = ExcelAccessor.getInstance(new FileInputStream(outputFile));
    documentDelivered = newExcelFileAccessor.parse(5, DocumentDelivered.class);
    
    assertEquals(Integer.valueOf(5), documentDelivered.getRowNum());
    assertEquals("MLMC", documentDelivered.getApplication());
    assertEquals("PIL", documentDelivered.getType());
    assertEquals("DOCi", documentDelivered.getTypeDoc());
    assertEquals("Manuel interne de livraison d’une version (génération, packaging, livraison)", documentDelivered.getName());
    assertEquals("", documentDelivered.getDetails());
    assertEquals("NPI_DCV_145", documentDelivered.getReference());
    assertEquals(Integer.valueOf(13), documentDelivered.getVersion());
  }
  
  /**
   * Lecture de l'ensemble des lignes du cartouche: 20 lignes datées. Ajout d'une ligne et relecture de tout.
   * 
   * @throws IOException
   */
  @Test
  public void lectureEcritureIntro_cartouche() throws IOException {
    
    // première lecture
    boolean readableLine = true;
    int rowNumber = 2;
    DateMidnight dateMinuit = new DateMidnight(2012, 1, 1);
    Cartouche cartoucheLine = excelAccessor.parse(rowNumber, Cartouche.class);
    while (readableLine) {
      assertEquals("userName" + (rowNumber - 1), cartoucheLine.getUserName());
      assertEquals(dateMinuit.plusDays(rowNumber - 2).toDate(), cartoucheLine.getDate());
      assertEquals("userAction" + (rowNumber - 1), cartoucheLine.getAction());
      assertEquals(BigDecimal.valueOf(rowNumber - 1).compareTo(cartoucheLine.getCost()), 0);
      cartoucheLine = excelAccessor.parse(++rowNumber, Cartouche.class);
      readableLine = cartoucheLine != null;
    }
    assertEquals(20, rowNumber - 1 - 1);
    LOGGER.debug("{} lignes lues dans le cartouche", rowNumber - 1 - 1);
    
    // ajout d'une ligne
    Cartouche cartoucheNouveau = excelAccessor.add(Cartouche.class);
    cartoucheNouveau.setUserName("userName" + (rowNumber - 1));
    cartoucheNouveau.setDate(dateMinuit.plusDays(rowNumber - 2).toDate());
    cartoucheNouveau.setAction("userAction" + (rowNumber - 1));
    cartoucheNouveau.setCost(BigDecimal.valueOf(rowNumber - 1));
    
    // sauvegarde du résultat
    File outputFile = File.createTempFile("test_cartouche_data", ".xls", new File("target"));
    FileOutputStream fileOutputStream = new FileOutputStream(outputFile);
    excelAccessor.write(fileOutputStream);
    fileOutputStream.close();
    
    // relecture de l'ensemble
    ExcelAccessor newExcelFileAccessor = ExcelAccessor.getInstance(new FileInputStream(outputFile));
    readableLine = true;
    rowNumber = 2;
    cartoucheLine = newExcelFileAccessor.parse(rowNumber, Cartouche.class);
    while (readableLine) {
      assertEquals("userName" + (rowNumber - 1), cartoucheLine.getUserName());
      assertEquals(dateMinuit.plusDays(rowNumber - 2).toDate(), cartoucheLine.getDate());
      assertEquals("userAction" + (rowNumber - 1), cartoucheLine.getAction());
      assertEquals(BigDecimal.valueOf(rowNumber - 1).compareTo(cartoucheLine.getCost()), 0);
      cartoucheLine = newExcelFileAccessor.parse(++rowNumber, Cartouche.class);
      readableLine = cartoucheLine != null;
    }
    assertEquals(21, rowNumber - 1 - 1);
    LOGGER.debug("{} lignes relues dans le cartouche", rowNumber - 1 - 1);
  }
}
