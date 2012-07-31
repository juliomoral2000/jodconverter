//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.task.ErrorCodeIOException;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;
import org.apache.commons.io.FilenameUtils;
import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.document.DocumentFormatRegistry;
import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeTask;
import org.artofsolving.jodconverter.stream.OOoInputStream;
import org.artofsolving.jodconverter.stream.OOoOutputStream;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

import static org.artofsolving.jodconverter.office.OfficeUtils.*;

public class StreamConversionTask implements OfficeTask {

    private String inputFullFileName;
    private String outputFullFileName;
    private DocumentFormatRegistry formatRegistry;
    private OOoInputStream oooInputStream;
    private OOoOutputStream oooOutputStream;
    private DocumentFormat inputFormat;
    private DocumentFormat outputFormat;
    private Map<String,?> defaultLoadProperties;

    public StreamConversionTask(String inputFullFileName, String outputFullFileName, DocumentFormatRegistry formatRegistry) {
        this.inputFullFileName = inputFullFileName;
        this.outputFullFileName = outputFullFileName;
        this.formatRegistry = formatRegistry;
        inputFormat = formatRegistry.getFormatByExtension(FilenameUtils.getExtension(inputFullFileName));
        outputFormat = formatRegistry.getFormatByExtension(FilenameUtils.getExtension(outputFullFileName));
        /*this.oooInputStream = oooInputStream;
        this.oooOutputStream = oooOutputStream;
        this.outputFormat = outputFormat;*/
    }

    public void setDefaultLoadProperties(Map<String, ?> defaultLoadProperties) {
        this.defaultLoadProperties = defaultLoadProperties;
    }



    protected Map<String,Object> getLoadProperties() {
        Map<String,Object> loadProperties = new HashMap<String,Object>();
        if (defaultLoadProperties != null) {
            loadProperties.putAll(defaultLoadProperties);
        }
        if (inputFormat != null && inputFormat.getLoadProperties() != null) {
            loadProperties.putAll(inputFormat.getLoadProperties());
        }
        return loadProperties;
    }

    protected Map<String,?> getStoreProperties(XComponent document) {
        DocumentFamily family = OfficeDocumentUtils.getDocumentFamily(document);
        return outputFormat.getStoreProperties(family);
    }

    public void execute(OfficeContext context) throws OfficeException {
        XComponent document = null;
        try {
            document = loadDocument(context);
            modifyDocument(document);
            storeDocument(document);
        } catch (OfficeException officeException) {
            throw officeException;
        } catch (Exception exception) {
            throw new OfficeException("conversion failed", exception);
        } finally {
            if (document != null) {
                XCloseable closeable = cast(XCloseable.class, document);
                if (closeable != null) {
                    try {
                        closeable.close(true);
                    } catch (CloseVetoException closeVetoException) {
                        // whoever raised the veto should close the document
                    }
                } else {
                    document.dispose();
                }
            }
        }
    }

    private XComponent loadDocument(OfficeContext context) throws OfficeException {
        createOOoInputStream();
        XComponentLoader loader = cast(XComponentLoader.class, context.getService(SERVICE_DESKTOP));
        Map<String,Object> loadProperties = getLoadProperties();
        loadProperties.put("InputStream", oooInputStream);
        XComponent document = null;
        try {
            document = loader.loadComponentFromURL("private:stream", "_blank", 0, toUnoProperties(loadProperties));
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new OfficeException("could not load document: " + inputFullFileName, illegalArgumentException);
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not load document: "  + inputFullFileName + "; errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);        } catch (IOException ioException) {
            throw new OfficeException("could not load document: "  + inputFullFileName, ioException);
        }
        if (document == null) {
            throw new OfficeException("could not load document: "  + inputFullFileName);
        }
        return document;
    }

    private void createOOoInputStream() {
        // Creamos un InputStream en buffer almacenando un OOoInputStream
        InputStream inputFile = null;
        try {
            inputFile = new BufferedInputStream(new FileInputStream(inputFullFileName));
            // Creamos un OutputStream de arreglo de bytes
            ByteArrayOutputStream bytes = new ByteArrayOutputStream();
            // creamos un arreglo de byte para escritura
            byte[] byteBuffer = new byte[4096];
            // inicializamos el contador del tamaño del bufer de escritura
            int byteBufferLength = 0;
            // escribimos en OutputStream cada 4096 bytes
            while ((byteBufferLength = inputFile.read(byteBuffer)) > 0) {
                bytes.write(byteBuffer,0,byteBufferLength);
            }
            inputFile.close();  // cerramos el InputStream
            // Creamos un nuevo OOoInputStream a partir del ByteArrayOutputStream
            oooInputStream = new OOoInputStream(bytes.toByteArray());
        } catch (FileNotFoundException e) {
            throw new OfficeException("Error reading input File "+inputFullFileName);
        } catch (java.io.IOException e) {
            throw new OfficeException("Error reading input File "+inputFullFileName);
        }
    }

    /**
     * Override to modify the document after it has been loaded and before it gets
     * saved in the new format.
     * <p>
     * Does nothing by default.
     *
     * @param document
     * @throws org.artofsolving.jodconverter.office.OfficeException
     */
    protected void modifyDocument(XComponent document) throws OfficeException {
    	// noop
    }

    private void storeDocument(XComponent document) throws OfficeException {
        createOOoOutputStream();
        Map<String,?> storeProperties = getStoreProperties(document);
        if (storeProperties == null) {
            throw new OfficeException("unsupported conversion");
        }
        String filterName = (String) storeProperties.get("FilterName");
        PropertyValue[] convrsProps = new PropertyValue[2];
        convrsProps[0] = new PropertyValue();
        convrsProps[1] = new PropertyValue();
        convrsProps[0].Name = "OutputStream";
        convrsProps[0].Value = oooOutputStream;
        convrsProps[1].Name = "FilterName";
        convrsProps[1].Value = filterName;
        try {
            cast(XStorable.class, document).storeToURL("private:stream", convrsProps);
            FileOutputStream outputFile = new FileOutputStream(outputFullFileName);
            outputFile.write(oooOutputStream.toByteArray());
            outputFile.close();
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not store document: " + outputFullFileName + "; errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not store document: " + outputFullFileName, ioException);
        } catch (FileNotFoundException fnfe) {
            throw new OfficeException("could not store document: " + outputFullFileName, fnfe);
        } catch (java.io.IOException ioe) {
            throw new OfficeException("could not store document: " + outputFullFileName, ioe);
        }
    }

    private void createOOoOutputStream() {
        oooOutputStream = new OOoOutputStream();
    }

}
