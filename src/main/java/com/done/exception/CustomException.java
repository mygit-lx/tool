package com.done.exception;


public class CustomException extends Exception {

    private String msg;

    public CustomException(String msg) {
        super();
        this.msg = msg;
    }

    public String getMsg() {
        return msg;
    }
}