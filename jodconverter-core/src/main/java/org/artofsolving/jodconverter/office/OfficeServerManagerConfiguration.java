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
package org.artofsolving.jodconverter.office;

/**
 * Clase de configuracion de conexion del Servidor OOo remoto
 */
public class OfficeServerManagerConfiguration {

    private int portNumber = 2002;
    private String hostIp = "127.0.0.1";
    private boolean connectOnStart = true;

    /**
     * constructor con valores por defecto
     */
    public void OfficeServerManagerConfiguration(){}

    /**
     *
     * @param hostIp
     * @param portNumber
     * @param connectOnStart
     */
    public OfficeServerManagerConfiguration(String hostIp, int portNumber, boolean connectOnStart) {
        setSocketParam(hostIp, portNumber);
        setConnectOnStart(connectOnStart);
    }

    public OfficeServerManagerConfiguration(String hostIp, int portNumber) {
        this(hostIp, portNumber, true);
    }

    protected void setSocketParam(String hostIp, int portNumber) {
        this.hostIp = hostIp;
        this.portNumber = portNumber;
    }
    protected void setConnectOnStart(boolean connectOnStart) {
        this.connectOnStart = connectOnStart;
    }

    public OfficeServerManager buildOfficeServerManager() {
        UnoUrl unoUrl = UnoUrl.socket(hostIp, portNumber);
        return new OfficeServerManager(unoUrl, connectOnStart);
    }

}
