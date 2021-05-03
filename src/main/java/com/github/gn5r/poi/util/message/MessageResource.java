package com.github.gn5r.poi.util.message;

public interface MessageResource {
    
    String getCode();

    String getMessagePattern();

    String getMessage(Object... args);

    String getSimpleMessage(Object... args);
}
