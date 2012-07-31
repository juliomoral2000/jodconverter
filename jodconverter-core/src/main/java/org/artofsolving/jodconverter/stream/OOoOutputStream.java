package org.artofsolving.jodconverter.stream;

import com.sun.star.io.BufferSizeExceededException;
import com.sun.star.io.NotConnectedException;
import com.sun.star.io.XOutputStream;

import java.io.ByteArrayOutputStream;

/**
 * <a href="http://www.oooforum.org/forum/viewtopic.phtml?t=13205">OOInputStream from the thread <b>OOo-Java: Using XInputStream...</b></a>
 * Otro link:</br>
 * <a href="http://user.services.openoffice.org/en/forum/viewtopic.php?t=3801">
 * [Java solution] Using XInputStream and XOutputStream</b>
 * </a>
 */
public class OOoOutputStream extends ByteArrayOutputStream implements XOutputStream {

    public OOoOutputStream() {
        super(32768);
    }


    //
    // Implement XOutputStream
    //
    public void writeBytes(byte[] values) throws NotConnectedException, BufferSizeExceededException, com.sun.star.io.IOException {
        try {
            this.write(values);
        } catch (java.io.IOException e) {
            throw (new com.sun.star.io.IOException(e.getMessage()));
        }
    }

    public void closeOutput() throws NotConnectedException, BufferSizeExceededException, com.sun.star.io.IOException {
        try {
            super.flush();
            super.close();
        } catch (java.io.IOException e) {
            throw (new com.sun.star.io.IOException(e.getMessage()));
        }
    }

    @Override
    public void flush() {
        try {
            super.flush();
        } catch (java.io.IOException e) {
        }
    }
}