package org.celper;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * default Delimiter ","
 * default prefix "\"\""
 * default suffix "\n"
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.SOURCE)
public @interface CSVConfig {
    String delimiter();
    String prefix();
    String suffix();
}
