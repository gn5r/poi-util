package com.github.gn5r.poi.util.message;

import java.text.MessageFormat;

public enum Message implements MessageResource {

    POI0001("\"{0}\" does not annotate the \"{1}\""),
    POI0002("can not find name \"{0}\" defined cell"),
    POI0003("the class of a \"{0}\" is a List. this method is List type field is not allowed"),
    POI0004("the class of a \"{0}\" is a single Object. this method is single Object type field is not allowed"),
    POI0005("\"{0}\" is type a \"{1}\"");

    private final String messgaePattern;

    private Message(String messagePattern) {
        this.messgaePattern = messagePattern;
    }

    @Override
    public String getCode() {
        return name();
    }

    @Override
    public String getMessagePattern() {
        return this.messgaePattern;
    }

    @Override
    public String getMessage(Object... args) {
        String simpleMessage = this.getSimpleMessageInternal(args);
        String code = name();
        return "[" + code + "] " + simpleMessage;
    }

    @Override
    public String getSimpleMessage(Object... args) {
        return null;
    }

    protected String getSimpleMessageInternal(Object... args) {
        return MessageFormat.format(this.messgaePattern, args);
    }
}
