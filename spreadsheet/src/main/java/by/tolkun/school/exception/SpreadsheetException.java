package by.tolkun.school.exception;

/**
 * Class to represent exception case.
 */
public class SpreadsheetException extends Exception {

    /**
     * Default constructor.
     */
    public SpreadsheetException() {
    }

    /**
     * Constructor with parameters.
     *
     * @param message the message of exception
     */
    public SpreadsheetException(String message) {
        super(message);
    }

    /**
     * Constructor with parameters.
     *
     * @param message the message of exception
     * @param cause   the cause (A {@code null} value is permitted, and
     *                indicates that the cause is nonexistent or unknown.)
     */
    public SpreadsheetException(String message, Throwable cause) {
        super(message, cause);
    }

    /**
     * Constructor with parameters.
     *
     * @param cause the cause (A {@code null} value is permitted, and indicates
     *              that the cause is nonexistent or unknown.)
     */
    public SpreadsheetException(Throwable cause) {
        super(cause);
    }

    /**
     * Constructor with parameters.
     *
     * @param message            the message of exception
     * @param cause              the cause. (A {@code null} value is permitted,
     *                           and indicates that the cause is nonexistent or
     *                           unknown.)
     * @param enableSuppression  whether or not suppression is enabled
     *                           or disabled
     * @param writableStackTrace whether or not the stack trace should be
     *                           writable
     */
    public SpreadsheetException(String message, Throwable cause,
                                boolean enableSuppression,
                                boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
