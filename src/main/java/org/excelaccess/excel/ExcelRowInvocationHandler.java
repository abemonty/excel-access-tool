package org.excelaccess.excel;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BLANK;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;
import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.Date;
import java.util.IllegalFormatException;
import java.util.Locale;
import org.apache.commons.lang.NotImplementedException;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.excelaccess.excel.model.annotation.ExcelCell;
import org.excelaccess.excel.model.annotation.ExcelCellFormat;
import org.excelaccess.excel.model.annotation.ExcelInternal;
import org.excelaccess.excel.model.annotation.RepeatableExcelCell;
import org.excelaccess.excel.utils.ExcelHandlerUtils;
import org.excelaccess.excel.utils.ExcelUtils;
import org.joda.time.DateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Accesseur des méthodes de manipulations d'une ligne Excel
 * 
 * @author Loic Abemonty
 */
public class ExcelRowInvocationHandler implements InvocationHandler {
  
  private static final String BLANK_STRING_VALUE = "";
  
  private static final Logger LOGGER = LoggerFactory.getLogger(ExcelRowInvocationHandler.class);
  
  /**
   * Création d'une cellule dans une ligne.
   * 
   * @param row
   *          non null
   * @param method
   *          non null et avec l'annotation {@link ExcelCell}
   * @param columnDelta
   *          décalage après la première cellule concernée
   * @return null s'il est impossible de créer la cellule
   */
  public static HSSFCell createCellFromMethod(HSSFRow row, Method method, int columnDelta) {
    HSSFCell cell = null;
    
    if (row == null | method == null) {
      LOGGER.warn("Impossible de créer la cellule voulue, la ligne n'existe pas ou la méthode d'application est incorrecte");
      return null;
    }
    ExcelCell excelCell = ExcelHandlerUtils.getAnnotation(ExcelCell.class, method);
    if (excelCell == null) {
      LOGGER.warn("Impossible de créer la cellule voulue, la méthode ne possède pas l'annotation ExcelCell (" + method.getName() + ")");
      return null;
    }
    
    // on crée la cellule avec le type qui nous intérèsse
    Class<?> returnType = method.getReturnType();
    if (void.class.equals(returnType)) {
      // probablement le cas d'un setter
      // nous allons nous baser sur le premier paramètre
      Class<?>[] parameterTypes = method.getParameterTypes();
      if (parameterTypes.length < 1) {
        // cas non géré
        LOGGER.warn("Impossible de créer une cellule à partir d'une cellule sans type de retour et avec des paramètres vides.");
        return null;
      }
      returnType = parameterTypes[0];
    }
    
    // FIXME loic - mieux gérer les types, genre "tous les numériques" ...
    // détermination du type de cellule
    int cellTypeId = CELL_TYPE_BLANK;
    if (String.class.equals(returnType)) {
      cellTypeId = CELL_TYPE_STRING;
    } else if (DateTime.class.equals(returnType) || Date.class.equals(returnType)) {
      // Loic : gestion native des dates, plus propre. (anciennement gestion des dates par un type String)
      cellTypeId = CELL_TYPE_NUMERIC;
    } else if (int.class.equals(returnType) || Integer.class.equals(returnType)) {
      cellTypeId = CELL_TYPE_NUMERIC;
    } else if (Double.class.equals(returnType)) {
      cellTypeId = CELL_TYPE_NUMERIC;
    } else if (Float.class.equals(returnType)) {
      cellTypeId = CELL_TYPE_NUMERIC;
    } else if (Long.class.equals(returnType)) {
      cellTypeId = CELL_TYPE_NUMERIC;
    } else if (BigDecimal.class.equals(returnType)) {
      cellTypeId = CELL_TYPE_NUMERIC;
    }
    
    int cellColumnId = getColumnIdForMethod(excelCell);
    if (cellTypeId != CELL_TYPE_BLANK) {
      cell = row.createCell(cellColumnId + columnDelta, cellTypeId);
      LOGGER.debug("création de la cellule " + cellColumnId + columnDelta + "(" + cellTypeId + ", " + returnType.getName() + ")");
    } else {
      LOGGER.warn("création de la cellule " + cellColumnId + columnDelta + " non effectuée, type non géré : " + method.getName() + " -> "
          + returnType.getName());
    }
    
    return cell;
  }
  
  /**
   * Détermine le numéro de la colonne d'une méthode. Gestion des lettres et des chiffres.
   * 
   * @param excelCell
   *          non null
   * @see ExcelCell
   * @return le numéro de la colonne, 0-based
   */
  private static int getColumnIdForMethod(ExcelCell excelCell) {
    int columnNumber = 0;
    if (excelCell != null) {
      columnNumber = excelCell.value();
      if (columnNumber == 0 && StringUtils.isNotBlank(excelCell.name())) {
        columnNumber = ExcelUtils.computeColumnIndexFromLetters(excelCell.name());
      }
    }
    return columnNumber;
  }
  
  /**
   * Détermine le numéro de la colonne d'une méthode. Gestion des lettres et des chiffres.
   * 
   * @see ExcelCell
   * @param method
   * @return le numéro de la colonne, 0-based
   */
  private static int getColumnIdForMethod(Method method) {
    ExcelCell excelCell = ExcelHandlerUtils.getAnnotation(ExcelCell.class, method);
    return getColumnIdForMethod(excelCell);
  }
  
  /**
   * vérifie qu'un nom de méthode est un nom de getter, en <tt>"getXxx()"</tt> , <tt>"hasXxx()"</tt> ou <tt>"isXxx()"</tt>, mais pas
   * <tt>"isNullXxx()"</tt>.
   */
  private static boolean isMethodNameGetter(final String methodName) {
    
    return methodName.startsWith("get") || methodName.startsWith("has")
        || (methodName.startsWith("is") && !methodName.startsWith("isNull"));
  }
  
  private static boolean isMethodNameSetter(final String methodName) {
    return methodName.startsWith("set");
  }
  
  private final HSSFRow row;
  
  private final Class<?> type;
  
  /**
   * Création du handler qui va s'occuper de gérer les méthodes du proxy pour une ligne particulière représentée par une classe spécifique.
   * 
   * @param clazz
   *          not null
   * @param row
   *          not null
   */
  public ExcelRowInvocationHandler(Class<?> clazz, HSSFRow row) {
    if (clazz == null || row == null) {
      throw new IllegalArgumentException("La classe représentant la ligne ou la ligne elle-même est null.");
    }
    this.type = clazz;
    this.row = row;
  }
  
  @Override
  public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
    
    final String methodName = method.getName();
    final boolean noArg = (args == null) || args.length == 0;
    
    // vérification de la présence d'une méthode interne
    ExcelInternal excelInternal = ExcelHandlerUtils.getAnnotation(ExcelInternal.class, method);
    if (excelInternal != null) {
      // accès à des données internes.
      return doInternalMethod(excelInternal);
    }
    
    if ("toString".equals(methodName) && noArg) {
      return doToString();
    } else if ("hashCode".equals(methodName) && noArg) {
      return doHashCode();
    } else if ("equals".equals(methodName) && !noArg && args.length == 1) {
      return doEquals(proxy, args[0]);
    } else if (isMethodNameGetter(methodName)) {
      if (noArg) {
        return doGet(method);
      } else {
        if (args.length == 1 && Integer.class.equals(args[0].getClass())) {
          // getter indexé
          return doGet(method, (Integer)args[0]);
        }
        // else return null
      }
    } else if (isMethodNameSetter(methodName) && !noArg) {
      if (args.length == 1) {
        // setter classique
        return doSet(method, args[0]);
      } else if (args.length == 2 && args[0] != null && Integer.class.equals(args[0].getClass())) {
        // setter indexé
        return doSet(method, (Integer)args[0], args[1]);
      }
      // else return null
    }
    
    return null;
  }
  
  private Object doEquals(Object proxy, Object object) {
    throw new NotImplementedException();
  }
  
  private Object doHashCode() {
    throw new NotImplementedException();
  }
  
  private Object doInternalMethod(ExcelInternal excelInternal) throws IllegalAccessException {
    
    switch (excelInternal.value()) {
      case ROW_LINE:
        return this.row.getRowNum();
      default:
        // aucune autre méthode n'est géré en ExcelInternal
        throw new IllegalAccessException("méthode interne inconnue (" + excelInternal.value() + ")");
    }
  }
  
  private Object doToString() {
    throw new NotImplementedException();
  }
  
  /**
   * /** Recherche d'une cellule d'après l'annotation d'une méthode.
   * 
   * @param method
   *          non null
   * @param columnDelta
   *          le décalage de l'index de la méthode
   * @return null si non trouvé
   * @see ExcelCell
   * @see RepeatableExcelCell
   */
  private HSSFCell getCellForMethod(Method method, int columnDelta) {
    int columnNumber = getColumnIdForMethod(method);
    HSSFCell cell = getRow().getCell(columnNumber + columnDelta);
    
    return cell;
  }
  
  /**
   * lit la valeur d'une cellule, en la convertissant dans le bon type. <br/>
   * <b>Il y a des problèmes de lecture d'un excel car des cases sont positionnées automatiquement par POI sur un type (double par exemple)
   * et on veut les lire sur un autre type (String).</b>
   * 
   * @param method
   *          la méthode <em>getter</em> correspondante.
   * @return la valeur lue.
   */
  protected Object doGet(Method method) {
    
    return doGet(method, -1);
    
  }
  
  /**
   * lit la valeur d'une cellule, en la convertissant dans le bon type. <br/>
   * <b>Il y a des problèmes de lecture d'un excel car des cases sont positionnées automatiquement par POI sur un type (double par exemple)
   * et on veut les lire sur un autre type (String).</b>
   * 
   * @param method
   *          la méthode <em>getter</em> correspondante.
   * @param index
   *          le décalage dans la lecteur de la méthode ; 0-based ; non considéré si négatif
   * @return la valeur lue.
   * @see RepeatableExcelCell
   */
  protected Object doGet(Method method, Integer index) {
    
    int columnDelta = 0;
    if (index >= 0) {
      RepeatableExcelCell repeatableExcelCell = ExcelHandlerUtils.getAnnotation(RepeatableExcelCell.class, method);
      if (repeatableExcelCell == null || repeatableExcelCell.size() <= index) {
        // si l'index est supérieur à zéro il faut cette annotation et que la taille de l'annotation soit
        // correcte
        return null;
      }
      columnDelta = index + (index * repeatableExcelCell.jump());
    }
    
    HSSFCell cell = getCellForMethod(method, columnDelta);
    
    if (cell == null) {
      return null;
    }
    
    // éventuel formatage de la cellule
    ExcelCellFormat excelCellFormatable = ExcelHandlerUtils.getAnnotation(ExcelCellFormat.class, method);
    Locale locale = null;
    NumberFormat numberFormat = null;
    if (excelCellFormatable != null && excelCellFormatable.localLanguage() != null) {
      locale = new Locale(excelCellFormatable.localLanguage());
    } else {
      locale = Locale.getDefault();
    }
    
    numberFormat = NumberFormat.getInstance(locale);
    
    int cellType = cell.getCellType();
    
    // Si la cellule est vide nous retournons vide
    if (CELL_TYPE_BLANK == cellType) {
      return null;
    }
    
    final Class<?> type = method.getReturnType();
    if (String.class.equals(type)) {
      if (CELL_TYPE_STRING == cellType) {
        return cell.getStringCellValue();
      } else {
        // numerique version, ne renvoie pas d'erreur si ce n'est pas une string
        return String.format("%d", Double.valueOf(cell.getNumericCellValue()).intValue());
      }
    }
    
    if (Date.class.equals(type)) {
      if (CELL_TYPE_STRING == cellType) {
        try {
          String stringCellValue = cell.getStringCellValue();
          if (StringUtils.isBlank(stringCellValue)) {
            return null;
          }
          return new DateTime(stringCellValue).toDate();
        } catch (IllegalArgumentException e) {
          return null;
        }
      } else {
        final Date date = cell.getDateCellValue();
        if (date == null) {
          return null;
        }
        return date;
      }
    }
    
    if (DateTime.class.equals(type) || Date.class.equals(type)) {
      if (CELL_TYPE_STRING == cellType) {
        try {
          String stringCellValue = cell.getStringCellValue();
          if (StringUtils.isBlank(stringCellValue)) {
            return null;
          }
          return new DateTime(stringCellValue);
        } catch (IllegalArgumentException e) {
          return null;
        }
      } else {
        final Date date = cell.getDateCellValue();
        if (date == null) {
          return null;
        }
        return new DateTime(date.getTime());
      }
    }
    
    /*
     * TODO (loic) : cette gestion des types numériques ne me plait pas, est-ce qu'un type primitif peut renvoyer null? Je ne pense pas.
     */
    if (int.class.equals(type) || Integer.class.equals(type)) {
      if (CELL_TYPE_STRING == cellType) {
        String value = cell.getStringCellValue();
        if (StringUtils.isNotBlank(value)) {
          return Integer.parseInt(value);
        }
        return null;
      } else {
        return Double.valueOf(cell.getNumericCellValue()).intValue();
      }
    }
    if (long.class.equals(type) || Long.class.equals(type)) {
      if (CELL_TYPE_STRING == cellType) {
        String value = cell.getStringCellValue();
        if (StringUtils.isNotBlank(value)) {
          try {
            return numberFormat.parse(value).longValue();
          } catch (ParseException e) {
            return null;
          }
        }
        return null;
      } else {
        return Double.valueOf(cell.getNumericCellValue()).longValue();
      }
    }
    
    if (double.class.equals(type) || Double.class.equals(type)) {
      if (CELL_TYPE_STRING == cellType) {
        String value = cell.getStringCellValue();
        if (StringUtils.isNotBlank(value)) {
          try {
            return numberFormat.parse(value).doubleValue();
          } catch (ParseException e) {
            return null;
          }
        }
        return null;
      } else {
        return Double.valueOf(cell.getNumericCellValue());
      }
    }
    
    if (float.class.equals(type) || Float.class.equals(type)) {
      if (CELL_TYPE_STRING == cellType) {
        String value = cell.getStringCellValue();
        if (StringUtils.isNotBlank(value)) {
          try {
            return numberFormat.parse(value).floatValue();
          } catch (ParseException e) {
            return null;
          }
        }
        return null;
      } else {
        return Double.valueOf(cell.getNumericCellValue()).floatValue();
      }
    }
    
    if (BigDecimal.class.equals(type)) {
      if (CELL_TYPE_STRING == cellType) {
        String value = cell.getStringCellValue();
        if (StringUtils.isNotBlank(value)) {
          try {
            return BigDecimal.valueOf(numberFormat.parse(value).longValue());
          } catch (ParseException e) {
            return null;
          }
        }
        return null;
      } else {
        return BigDecimal.valueOf(cell.getNumericCellValue());
      }
    }
    
    throw new NotImplementedException("type: " + type.getName());
  }
  
  /**
   * Gestion des getters indexés (avec un index) d'après la position de départ.
   * 
   * @param method
   *          la méthode non null
   * @param index
   *          l'index ; 0-based ; non considéré si négatif
   * @param object
   *          l'objet à setter
   * @return la cellule modifiée, null si rien ne s'est passé (cellule non trouvé)
   */
  protected Object doSet(Method method, Integer index, Object object) {
    Object value = BLANK_STRING_VALUE;
    
    value = object;
    
    int columnDelta = 0;
    if (index >= 0) {
      RepeatableExcelCell repeatableExcelCell = ExcelHandlerUtils.getAnnotation(RepeatableExcelCell.class, method);
      if (repeatableExcelCell == null || repeatableExcelCell.size() <= index) {
        // si l'index est supérieur à zéro il faut cette annotation et que la taille de l'annotation soit
        // correcte
        return null;
      }
      columnDelta = index + (index * repeatableExcelCell.jump());
    }
    
    // ExcelCell excelCell = ExcelHandlerUtils.getCellAnnotation(method);
    // int columnNumber = excelCell.value();
    // HSSFCell cell = this.row.getCell(columnNumber + columnDelta);
    
    HSSFCell cell = getCellForMethod(method, columnDelta);
    
    if (cell == null) {
      cell = createCellFromMethod(this.row, method, columnDelta);
      if (cell == null) {
        // impossible de créer la cellule
        return null;
      }
    }
    
    // éventuel formatage de la cellule avant écriture
    ExcelCellFormat excelCellFormatable = ExcelHandlerUtils.getAnnotation(ExcelCellFormat.class, method);
    if (excelCellFormatable != null) {
      if (value != null) {
        // impossible de formater une valeur null
        String stringFormat = excelCellFormatable.outputFormat();
        short excelFormat = excelCellFormatable.outputExcelFormat();
        // Formatage text par String.format
        if (StringUtils.isNotBlank(stringFormat)) {
          try {
            value = String.format(new Locale(excelCellFormatable.localLanguage()), stringFormat, value);
          } catch (final IllegalFormatException e) {
            LOGGER.warn("Erreur de format de la donnée de la ligne {}, colonne {} avec le format {}. La méthode concernée est {}",
                new Object[] {getRow().getRowNum(), index, excelCellFormatable.outputFormat(), method});
            return null;
          }
        }
        // ou formatage style Excel par l'API
        else {
          if (excelFormat >= 0) {
            cell.getCellStyle().setDataFormat(excelFormat);
          }
        }
      }
    }
    
    if (value == null) {
      // pas de set de null => set à ""
      value = BLANK_STRING_VALUE;
    }
    
    if (value instanceof String) {
      cell.setCellValue((String)value);
    } else if (value instanceof Integer) {
      cell.setCellValue((Integer)value);
    } else if (value instanceof Double) {
      cell.setCellValue((Double)value);
    } else if (value instanceof Float) {
      cell.setCellValue((Float)value);
    } else if (value instanceof Long) {
      cell.setCellValue((Long)value);
    } else if (value instanceof BigDecimal) {
      cell.setCellValue(((BigDecimal)value).floatValue());
    } else if (value instanceof Date) {
      cell.setCellValue((Date)value);
    } else if (value instanceof DateTime) {
      cell.setCellValue(((DateTime)value).toDate());
    } else {
      cell.setCellValue(value.toString());
    }
    
    return cell;
  }
  
  /**
   * Gestion des setters sur l'objet. <br/>
   * TODO : éligible à l'abstractibilité
   * 
   * @param method
   *          non null
   * @param object
   *          si null remplacé par "" (chaine vide)
   * @return null si le set ne s'est pas bien passé, la cellule concernée sinon
   */
  protected Object doSet(Method method, Object object) {
    
    return doSet(method, -1, object);
  }
  
  protected HSSFRow getRow() {
    return row;
  }
  
  // /**
  // * Recherche d'une cellule d'après l'annotation d'une méthode.
  // *
  // * @param method
  // * non null
  // * @return null si non trouvé
  // *
  // * @see ExcelCell
  // */
  // private HSSFCell getCellForMethod(Method method) {
  //
  // return getCellForMethod(method, 0);
  // }
  
  protected Class<?> getType() {
    return type;
  }
}
