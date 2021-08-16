package org.svetikov.reportlaboratorymdf;

import org.springframework.core.annotation.AliasFor;

import java.lang.annotation.*;

@Target({ElementType.PARAMETER})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface NameList {
//    @AliasFor("name")
    String value() default "";
}
