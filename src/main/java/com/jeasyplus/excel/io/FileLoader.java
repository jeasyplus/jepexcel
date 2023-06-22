package com.jeasyplus.excel.io;

import java.io.*;

public class FileLoader {

    public static InputStream getInputStream(String path) throws IOException {
        try (FileInputStream input = new FileInputStream(path)) {
            byte[] buffer  = new byte[input.available()];
            input.read(buffer);
            return new ByteArrayInputStream(buffer);
        }
    }

}
