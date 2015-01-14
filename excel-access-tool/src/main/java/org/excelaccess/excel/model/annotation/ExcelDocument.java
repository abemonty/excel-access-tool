package org.excelaccess.excel.model.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * utiliser cette annotation sur une interface de <em>getters</em>, pour indiquer à quel numéro de feuille du document
 * Excel elle correspond, et à quel numéro de ligne il faut commencer le parsing.
 * 
 * @author David Andrianavalontsalama
 * @author Loic Abemonty
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelDocument {

    /**
     * Nom de l'onglet. Obligatoire.
     * 
     * @return le nom de l'onglet
     * */
    String sheetName();

    /**
     * numéro de la première ligne, 0-based. Obligatoire.
     * 
     * @return le numéro de la ligne
     */
    int startAtRow();

}
