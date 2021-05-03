package com.github.gn5r.poi.util.exception;

import com.github.gn5r.poi.util.message.MessageResource;

/**
 * 
 * @author gn5r
 */
public class PoiIllegalArgumentException extends PoiException {

    private final String parameterName;

    public PoiIllegalArgumentException(MessageResource messageResource, String parameterName) {
        super(messageResource, parameterName);
        this.parameterName = parameterName;
    }

    public String getParameterName() {
        return this.parameterName;
    }
}
