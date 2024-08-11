package org.celper;

import org.celper.core2.style.SheetLayoutConfigurer;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.SOURCE)
public @interface SheetLayout {
    Class<? extends SheetLayoutConfigurer> value();
}
