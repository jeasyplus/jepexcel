package com.jeasyplus.excel.io;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class FileWriter {

    public static OutputStream getOutStream(String path) throws FileNotFoundException {
        return new FileOutputStream(path);
    }

}
