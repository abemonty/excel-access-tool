package org.excelaccess.excel.utils;

/**
 * Classe utilitaire de manipulation d'objet Excel.
 * 
 * @author Loic Abemonty
 * 
 */
public class ExcelUtils {

    /**
     * Calcule un numéro de colonne excel à partir de sa représentation textuelle en lettres.
     * 
     * @param columnLettersIndex
     * @return une valeur négative si un problème apparait.
     */
    public static int computeColumnIndexFromLetters(String columnLettersIndex) {
        int valueA = Character.getNumericValue('A');
        int baseAlpha = 26;
        // int i = 0;
        double computedIndex = 0;

        char[] charArray = columnLettersIndex.toUpperCase().toCharArray();
        for (int i = 0; i < charArray.length; i++) {
            int numericValue = Character.getNumericValue(charArray[i]) - valueA + 1;
            computedIndex = numericValue * Math.pow(baseAlpha, charArray.length - 1 - i) + computedIndex;
            // computedIndex = (computedIndex * (baseAlpha ^ i)) + (numericValue - valueA + 1);
            // i++; // décalage de base
        }
        return (int) (computedIndex - 1);
    }

}
