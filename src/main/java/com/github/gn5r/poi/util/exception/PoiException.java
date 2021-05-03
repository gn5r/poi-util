package com.github.gn5r.poi.util.exception;

import com.github.gn5r.poi.util.message.MessageResource;

public class PoiException extends RuntimeException {

    private static final long serialVersionUID = 1L;

    protected final MessageResource messageResource;
    protected final Object[] args;

    public PoiException(MessageResource messageResource, Object... args) {
        this(messageResource, null, args);
    }

    public PoiException(MessageResource messageResource, Throwable cause, Object... args) {
        super(messageResource.getMessage(args), cause);
        this.messageResource = messageResource;
        this.args = args;
    }

    public MessageResource getMessageResource() {
        return this.messageResource;
    }

    public Object getArgs() {
        return this.args;
    }
}
