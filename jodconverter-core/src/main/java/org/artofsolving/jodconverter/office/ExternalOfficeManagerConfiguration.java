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

public class ExternalOfficeManagerConfiguration {

    private OfficeConnectionProtocol connectionProtocol = OfficeConnectionProtocol.SOCKET;
    private int portNumber = 2002;
    private String hostIp = "127.0.0.1";
    private String pipeName = "office";
    private boolean connectOnStart = true;

    public ExternalOfficeManagerConfiguration setConnectionProtocol(OfficeConnectionProtocol connectionProtocol) {
        this.connectionProtocol = connectionProtocol;
        return this;
    }

    public ExternalOfficeManagerConfiguration setPortNumber(int portNumber) {
        this.portNumber = portNumber;
        return this;
    }

    public ExternalOfficeManagerConfiguration setPipeName(String pipeName) {
        this.pipeName = pipeName;
        return this;
    }

    public ExternalOfficeManagerConfiguration setHostIp(String hostIp) {
        this.hostIp = hostIp;
        return this;
    }

    public ExternalOfficeManagerConfiguration setSocketParam(String hostIp, int portNumber) {
        this.hostIp = hostIp;
        this.portNumber = portNumber;
        return this;
    }
    public ExternalOfficeManagerConfiguration setConnectOnStart(boolean connectOnStart) {
        this.connectOnStart = connectOnStart;
        return this;
    }

    public OfficeManager buildOfficeManager() {
        //UnoUrl unoUrl = connectionProtocol == OfficeConnectionProtocol.SOCKET ? UnoUrl.socket(portNumber) : UnoUrl.pipe(pipeName);
        UnoUrl unoUrl = connectionProtocol == OfficeConnectionProtocol.SOCKET ? UnoUrl.socket(hostIp, portNumber) : UnoUrl.pipe(pipeName);
        return new ExternalOfficeManager(unoUrl, connectOnStart);
    }

}
