package com.github.gn5r.poi.util.exception;

import com.github.gn5r.poi.util.annotation.Cell;
import com.github.gn5r.poi.util.message.Message;

/**
 * {@link Cell} がアノテートされていない場合にthrowされる
 * 
 * @author gn5r
 */
public class PoiNoAnnotatedException extends PoiException {

    private final String className;

    private final String annotationName;

    public PoiNoAnnotatedException(String className, String annotationName) {
        super(Message.POI0001, className, annotationName);
        this.className = className;
        this.annotationName = annotationName;
    }

    /**
     * @return the className
     */
    public String getClassName() {
        return className;
    }

    /**
     * @return the annotationName
     */
    public String getAnnotationName() {
        return annotationName;
    }
}
