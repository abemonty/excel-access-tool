package org.excelaccess.excel.model.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation pour accèder aux valeurs internes d'un objet : déclencher une méthode sur un objet interne.
 * 
 * @author Loic Abemonty
 * 
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelInternal {

    /**
     * Valeurs possibles : {@link ExcelInternalEnum#ROW}.
     */
    ExcelInternalEnum value();
}
