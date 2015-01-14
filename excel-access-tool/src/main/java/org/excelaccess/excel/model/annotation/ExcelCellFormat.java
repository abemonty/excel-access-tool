package org.excelaccess.excel.model.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * utiliser cette annotation sur une méthode <tt>getXxx()</tt> d'une interface de <em>getter</em>, pour indiquer quel
 * format de cellule utiliser.<br/>
 * Utilisation des mêmes paramètres de formatage que {@link String#format(String, Object...)}.<br/>
 * Exemple pour formater une date dans le format "YYYY-mm-DD" : <code>"%1$tY-%&lt;tm-%&lt;td"</code>.
 * 
 * @author David Andrianavalontsalama
 * @author Loic Abemonty
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellFormat {

    /**
     * local Language, correspond à la langue locale
     * 
     * @see java.util.Locale
     * @exemple valeur possible java.util.Locale.FRENCH.getLanguage()
     * @defaut "fr"
     * @return the string
     */
    String localLanguage() default "fr";

    /**
     * Formatage de la cellule d'après un format prédéfini dans Excel. A utiliser à la place de {@link #outputFormat()}
     * ; si les deux sont définis, c'est {@link #outputFormat()} qui prend la priorité. <br/>
     * Valeur par défaut : -1.
     * <p/>
     * <p/>
     * 0, "General"<br/>
     * 1, "0"<br/>
     * 2, "0.00"<br/>
     * 3, "#,##0"<br/>
     * 4, "#,##0.00"<br/>
     * 5, "$#,##0_);($#,##0)"<br/>
     * 6, "$#,##0_);[Red]($#,##0)"<br/>
     * 7, "$#,##0.00);($#,##0.00)"<br/>
     * 8, "$#,##0.00_);[Red]($#,##0.00)"<br/>
     * 9, "0%"<br/>
     * 0xa, "0.00%"<br/>
     * 0xb, "0.00E+00"<br/>
     * 0xc, "# ?/?"<br/>
     * 0xd, "# ??/??"<br/>
     * 0xe, "m/d/yy"<br/>
     * 0xf, "d-mmm-yy"<br/>
     * 0x10, "d-mmm"<br/>
     * 0x11, "mmm-yy"<br/>
     * 0x12, "h:mm AM/PM"<br/>
     * 0x13, "h:mm:ss AM/PM"<br/>
     * 0x14, "h:mm"<br/>
     * 0x15, "h:mm:ss"<br/>
     * 0x16, "m/d/yy h:mm"<br/>
     * <p/>
     * // 0x17 - 0x24 reserved for international and undocumented 0x25, "#,##0_);(#,##0)"<br/>
     * 0x26, "#,##0_);[Red](#,##0)"<br/>
     * 0x27, "#,##0.00_);(#,##0.00)"<br/>
     * 0x28, "#,##0.00_);[Red](#,##0.00)"<br/>
     * 0x29, "_(*#,##0_);_(*(#,##0);_(* \"-\"_);_(@_)"<br/>
     * 0x2a, "_($*#,##0_);_($*(#,##0);_($* \"-\"_);_(@_)"<br/>
     * 0x2b, "_(*#,##0.00_);_(*(#,##0.00);_(*\"-\"??_);_(@_)"<br/>
     * 0x2c, "_($*#,##0.00_);_($*(#,##0.00);_($*\"-\"??_);_(@_)"<br/>
     * 0x2d, "mm:ss"<br/>
     * 0x2e, "[h]:mm:ss"<br/>
     * 0x2f, "mm:ss.0"<br/>
     * 0x30, "##0.0E+0"<br/>
     * 0x31, "@" - This is text format.<br/>
     * 0x31 "text" - Alias for "@"<br/>
     * <p/>
     * 
     */
    short outputExcelFormat() default -1;

    /**
     * Format d'export de la valeur dans la cellule. A utiliser à la place de {@link #outputExcelFormat()} ; si les deux
     * sont définis, c'est {@link #outputFormat()} qui prend la priorité.<br/>
     * Valeur par défaut : "".
     * 
     * @return the string
     */
    String outputFormat() default "";
}
