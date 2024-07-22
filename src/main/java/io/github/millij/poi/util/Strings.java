package io.github.millij.poi.util;

import java.util.Objects;


/**
 * Odd ball String Utilities
 * 
 * @since 3.1.0
 */
public final class Strings {

    private Strings() {
        super();
        // Utility Class
    }


    // Util Methods
    // ------------------------------------------------------------------------

    /**
     * Normalize the string by removing unwanted characters from the String.
     * 
     * @param inStr Input String
     * @param replacement Replacement character for the unwanted characters
     * 
     * @return Clean / Normalized String
     */
    public static String normalize(final String inStr, final String replacement) {
        // Sanity checks
        if (Objects.isNull(inStr)) {
            return "";
        }

        // Special characters
        final String cleanStr = inStr.replaceAll("–", " ").replaceAll("[-\\[\\]/{}:.,;#%=()*+?\\^$|<>&\"\'\\\\]", " ");
        final String normalizedStr = cleanStr.toLowerCase().trim().replaceAll("\\s+", replacement);

        return normalizedStr;
    }

    /**
     * Normalize the string by replacing unwanted characters from the String with "_".
     * 
     * @param inStr Input String
     * 
     * @return Clean / Normalized String
     */
    public static String normalize(final String inStr) {
        return normalize(inStr, "_");
    }


}
