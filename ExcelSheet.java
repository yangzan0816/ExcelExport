package com.wxstore.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {
    String name() default "";
    //start todo：添加表格主题，本次并不实现该功能
    String sheetTitleName() default "";
    boolean isTitleRowMerge() default false;
    //end
}
