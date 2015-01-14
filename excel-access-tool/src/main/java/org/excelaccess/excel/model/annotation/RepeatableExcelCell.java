package org.excelaccess.excel.model.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * utiliser cette annotation sur une méthode <tt>getXxx()</tt> d'une interface
 * de <em>getter</em>, pour indiquer que cette cellule se répète {@link #size()}
 * fois et qu'il faut considérer un décalage de {@link #jump()} éléments.
 * <p/>
 * Une méthode utilisant cette méthode doit posséder en premier élément un
 * paramètre entier (int) ; c'est obligatoire pour le getter et pour le setter.
 * <br/>
 * <b>Le premier élément est à la position 0.</b>
 * 
 * @author Loic Abemonty
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface RepeatableExcelCell {

	/**
	 * Nombre d'élément à sauter avant la prochaine répétition ; 0 = pas de
	 * répétition;
	 * 
	 * @return 0 par défaut
	 */
	int jump() default 0;

	/**
	 * Nombre de répétition de la cellule.
	 * 
	 * @return 1 par defaut
	 */
	int size() default 1;
}
