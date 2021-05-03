package com.github.gn5r.poi.util.message;

public class MessageUtil {
    
    public static String getMessage(MessageResource messageResource, Object... args) {
        return messageResource.getMessage(args);
    }
}
