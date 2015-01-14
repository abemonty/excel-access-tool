package org.excelaccess.excel.model.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * utiliser cette annotation sur une méthode <tt>getXxx()</tt> d'une interface de <em>getter</em>, pour indiquer à quel
 * numéro de cellule correspond la valeur.
 * 
 * @author David Andrianavalontsalama
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {

    /**
     * Le nom de la colonne, "A", "B", ..., "AA", "AB", ...
     */
    String name() default "A";

    /**
     * Le numéro de la colonne, 0-based.
     */
    int value() default 0;
}
