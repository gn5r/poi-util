package com.github.gn5r.poi.util.logger;

import org.apache.commons.lang3.exception.ExceptionUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class PoiUtilLogger {
    
    private static final Logger log = LoggerFactory.getLogger(PoiUtilLogger.class);

    public static void info(String message) {
        log.info(message);
    }

    public static void debug(String message) {
        log.debug(message);
    }

    public static void warn(String message) {
        log.warn(message);
    }

    public static void error(Throwable cause) {
        log.error(ExceptionUtils.getStackTrace(cause));
    }
}
