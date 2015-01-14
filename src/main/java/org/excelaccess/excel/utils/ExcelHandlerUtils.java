package org.excelaccess.excel.utils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Method;

import org.apache.commons.lang.ArrayUtils;
import org.excelaccess.excel.model.annotation.ExcelCell;

/**
 * Méthodes utilitaires pour accèder aux cellules.
 * 
 * @author Loic Abemonty
 * 
 */
public class ExcelHandlerUtils {

    /**
     * les préfixes de méthodes reconnus.
     */
    public static final String[] METHOD_PREFIXES = { "set", "get", "has", "is", "add", "isNull" };

    /**
     * Recherche les informations Excel qui sont normalement sur le getter.
     * 
     * @param method
     *            non null
     * @return null si non trouvé
     */
    public static <T extends Annotation> T getAnnotation(Class<T> annotationClazz, Method method) {
        T excelCell = method.getAnnotation(annotationClazz);
        String methodName = method.getName();

        if (excelCell == null && methodName.startsWith("get") == false) {
            // recherchons le getter qui doit avoir les informations.
            for (final String prefix : METHOD_PREFIXES) {
                if (methodName.startsWith(prefix)) {
                    final String methodSuffix = methodName.substring(prefix.length());
                    final Method getterMethod;
                    try {
                        // supposons que dans l'ensemble des parametres, le dernier soit la valeur : donc à exclure du
                        // get
                        Class<?>[] parameterTypes = method.getParameterTypes();
                        if (parameterTypes.length > 0) {
                            parameterTypes = (Class<?>[]) ArrayUtils.subarray(parameterTypes, 0,
                                    parameterTypes.length - 1);
                        }
                        getterMethod = method.getDeclaringClass().getMethod("get" + methodSuffix, parameterTypes);
                    } catch (NoSuchMethodException e) {
                        return null;
                    }
                    return getAnnotation(annotationClazz, getterMethod);
                }
            }
        }
        return excelCell;
    }

    /**
     * Recherche les informations Excel qui sont normalement sur le getter, remonte l'arbre d'héritage pour trouver le
     * bon getter.
     * 
     * @param method
     *            non null
     * @return null si non trouvé
     */
    public static ExcelCell getCellAnnotation(Method method) {
        ExcelCell excelCell = method.getAnnotation(ExcelCell.class);
        String methodName = method.getName();

        if (excelCell == null && methodName.startsWith("get") == false) {
            // recherchons le getter qui doit avoir les informations.
            for (final String prefix : METHOD_PREFIXES) {
                if (methodName.startsWith(prefix)) {
                    final String methodSuffix = methodName.substring(prefix.length());
                    final Method getterMethod;
                    try {
                        getterMethod = method.getDeclaringClass().getMethod("get" + methodSuffix);
                    } catch (NoSuchMethodException e) {
                        return null;
                    }
                    return getCellAnnotation(getterMethod);
                }
            }
        }
        return excelCell;
    }
}