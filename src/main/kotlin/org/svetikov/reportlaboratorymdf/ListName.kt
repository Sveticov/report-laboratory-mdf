package org.svetikov.reportlaboratorymdf


@Target(AnnotationTarget.LOCAL_VARIABLE,AnnotationTarget.PROPERTY,AnnotationTarget.TYPE, AnnotationTarget.CLASS)//AnnotationTarget.LOCAL_VARIABLE, AnnotationTarget.PROPERTY
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ListName(val name:String)
