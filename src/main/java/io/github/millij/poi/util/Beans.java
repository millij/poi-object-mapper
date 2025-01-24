package io.github.millij.poi.util;

import static java.util.Calendar.HOUR_OF_DAY;
import static java.util.Calendar.MINUTE;
import static java.util.Calendar.SECOND;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Objects;

import org.apache.commons.beanutils.ConvertUtilsBean;
import org.apache.commons.beanutils.Converter;
import org.apache.commons.beanutils.PropertyUtilsBean;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.poi.ss.model.DateTimeType;


/**
 * Added Bean Utilities that are not directly available in the Apache BeanUtils library.
 */
public final class Beans {

    private static final Logger LOGGER = LoggerFactory.getLogger(Beans.class);

    private Beans() {
        super();
        // Utility Class
    }


    //
    // Constants

    private static final PropertyUtilsBean PROP_UTILS_BEAN = new PropertyUtilsBean();
    private static final ConvertUtilsBean CONVERT_UTILS_BEAN = new ConvertUtilsBean();


    // Static Utilities
    // ------------------------------------------------------------------------

    /**
     * Extracts the name of the field from its accessor method.
     * 
     * @param method any accessor {@link Method} of the field.
     * 
     * @return the name of the field.
     */
    public static String getFieldName(final Method method) {
        final String methodName = method.getName();
        return Introspector.decapitalize(methodName.substring(methodName.startsWith("is") ? 2 : 3));
    }


    /**
     * Given a Bean and a field of it, returns the value of the field converted to String.
     * 
     * <ul>
     * <li><code>null</code> is returned if the value of the field itself is null.</li>
     * <li>In the case of an Object type, its String representation will be returned.</li>
     * </ul>
     * 
     * @param beanObj bean of which the field value to be extracted.
     * @param fieldName Name of the property/field of the object.
     * 
     * @return the field value converted to String.
     * 
     * @throws Exception if the bean or the fields accessor methods are not accessible.
     */
    public static String getFieldValueAsString(final Object beanObj, final String fieldName) throws Exception {
        // Property Descriptor
        final PropertyDescriptor pd = PROP_UTILS_BEAN.getPropertyDescriptor(beanObj, fieldName);
        final Method getterMtd = pd.getReadMethod();

        final Object value = getterMtd.invoke(beanObj);
        final String cellValue = Objects.nonNull(value) ? String.valueOf(value) : null;
        return cellValue;
    }


    /**
     * Check whether a class is instantiable of not.
     * 
     * @param clz the {@link Class} which needs to verified.
     * 
     * @return false if the class in primitive/abstract/interface/array
     */
    public static boolean isInstantiableType(final Class<?> clz) {
        // Sanity checks
        if (Objects.isNull(clz)) {
            return false;
        }

        final int modifiers = clz.getModifiers();
        LOGGER.debug("Modifiers of Class : {} - {}", clz, modifiers);

        // Primitive / Abstract / Interface / Array
        if (clz.isPrimitive() || Modifier.isAbstract(modifiers) || clz.isInterface() || clz.isArray()) {
            return false;
        }

        return true;
    }


    // Bean Property :: Get
    // ------------------------------------------------------------------------

    public static Object getProperty(final Object bean, final String propName) throws Exception {
        final Object value = PROP_UTILS_BEAN.getSimpleProperty(bean, propName);
        return value;
    }


    // Bean Property :: Set
    // ------------------------------------------------------------------------

    /**
     */
    public static void setProperty(final Object target, final String propName, final Class<?> propType,
            final Object propValue) throws Exception {
        // Sanity checks
        if (Objects.isNull(propValue)) {
            return; // Skip Setter if property value is NULL
        }

        try {
            // Convert the specified value to the required type
            final Object newValue;
            if (propValue instanceof String) {
                newValue = CONVERT_UTILS_BEAN.convert((String) propValue, propType);
            } else {
                final Converter converter = CONVERT_UTILS_BEAN.lookup(propType);
                if (converter != null) {
                    newValue = converter.convert(propType, propValue);
                } else {
                    newValue = propValue;
                }
            }

            // Invoke the setter method
            PROP_UTILS_BEAN.setProperty(target, propName, newValue);

        } catch (Exception ex) {
            //
        }
    }


    /**
     */
    @SuppressWarnings({"unchecked", "rawtypes"})
    public static void setProperty(final Object target, final String propName, final Object propValue,
            final String format, final DateTimeType dateTimeType) throws Exception {
        // Sanity checks
        if (Objects.isNull(propValue)) {
            return; // Skip Setter if property value is NULL
        }

        // Calculate the property type
        final PropertyDescriptor descriptor = PROP_UTILS_BEAN.getPropertyDescriptor(target, propName);
        if (Objects.isNull(descriptor) || Objects.isNull(descriptor.getWriteMethod())) {
            return; // Skip this property setter
        }

        // Property Type
        final Class<?> propType = descriptor.getPropertyType();

        //
        // Handle Date/Time/Duration Cases
        if (!DateTimeType.NONE.equals(dateTimeType) || propType.equals(Date.class)) {
            setDateTimeProperty(target, propName, propType, propValue, format, dateTimeType);
            return;
        }

        //
        // Handle ENUM
        if (propType.isEnum() && (propValue instanceof String)) {
            final String cleanEnumStr = Strings.normalize((String) propValue).toUpperCase();
            final Enum<?> enumValue = Enum.valueOf((Class<? extends Enum>) propType, cleanEnumStr);
            setProperty(target, propName, propType, enumValue);
            return;
        }

        //
        // Handle Boolean
        if (propType.equals(Boolean.class) && (propValue instanceof String)) {
            // Cleanup Boolean String
            final String cleanBoolStr = Strings.normalize((String) propValue); // for cases like "FALSE()", "TRUE()"
            setProperty(target, propName, propType, cleanBoolStr);
            return;
        }

        //
        // Default Handling (for all other Types)
        setProperty(target, propName, propType, propValue);

    }


    /**
     * Set the Date/Time property of the Target Bean.
     */
    private static void setDateTimeProperty(final Object target, final String propName, final Class<?> propType,
            final Object propValue, final String format, final DateTimeType dateTimeType) throws Exception {
        // Input value Format
        final String dateFormatStr = Strings.isBlank(format) ? "dd/MM/yyyy" : format;

        // Parse
        final SimpleDateFormat dateFmt = new SimpleDateFormat(dateFormatStr);
        final Date dateValue = dateFmt.parse((String) propValue);

        // Check if the PropType is Date
        if (propType.equals(Date.class)) {
            setProperty(target, propName, propType, dateValue);

        } else {
            // Convert to Long
            final Long longValue;
            if (DateTimeType.DURATION.equals(dateTimeType)) {
                final Calendar calendar = Calendar.getInstance();
                calendar.setTime(dateValue);
                final long durationMillis = calendar.get(HOUR_OF_DAY) * 3600_000 + //
                        calendar.get(MINUTE) * 60_000 + //
                        calendar.get(SECOND) * 1_000;

                longValue = durationMillis;
            } else {
                longValue = dateValue.getTime();
            }

            setProperty(target, propName, propType, longValue);
        }

        //
    }

}
