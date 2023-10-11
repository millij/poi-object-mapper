package io.github.millij.poi.util;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * Added Bean Utilities that are not directly available in the Apache BeanUtils library.
 */
public final class Beans {

    private static final Logger LOGGER = LoggerFactory.getLogger(Spreadsheet.class);

    private Beans() {
        // Utility Class
    }


    // Static Utilities
    // ------------------------------------------------------------------------

    /**
     * Extrats the name of the field from its accessor method.
     * 
     * @param method any accessor {@link Method} of the field.
     * @return the name of the field.
     */
    public static String getFieldName(Method method) {
        String methodName = method.getName();
        return Introspector.decapitalize(methodName.substring(methodName.startsWith("is") ? 2 : 3));
    }


    /**
     * Given a Bean and a field of it, returns the value of the field converted to String.
     * 
     * <ul>
     * <li> <code>null</code> is returned if the value of the field itself is null.</li> 
     * <li> In the case of an Object type, its String representation will be returned.</li>
     * </ul>
     * 
     * @param beanObj bean of which the field value to be extracted.
     * @param fieldName Name of the property/field of the object.
     * @return the field value converted to String.
     * 
     * @throws Exception if the bean or the fields accessor methods are not accessible.
     */
    public static String getFieldValueAsString(Object beanObj, String fieldName) throws Exception {
        // Property Descriptor
        PropertyDescriptor pd = new PropertyDescriptor(fieldName, beanObj.getClass());
        Method getterMtd = pd.getReadMethod();

        Object value = getterMtd.invoke(beanObj);
        String cellValue = value != null ? String.valueOf(value) : null;

        return cellValue;
    }


    /**
     * Check whether a class is instantiable of not.
     * 
     * @param clz the {@link Class} which needs to verified.
     * @return false if the class in primitive/abstract/interface/array
     */
    public static boolean isInstantiableType(Class<?> clz) {
        // Sanity checks
        if (clz == null) {
            return false;
        }

        int modifiers = clz.getModifiers();
        LOGGER.debug("Modifiers of Class : {} - {}", clz, modifiers);

        // Primitive / Abstract / Interface / Array
        if (clz.isPrimitive() || Modifier.isAbstract(modifiers) || clz.isInterface() || clz.isArray()) {
            return false;
        }

        return true;
    }


}
