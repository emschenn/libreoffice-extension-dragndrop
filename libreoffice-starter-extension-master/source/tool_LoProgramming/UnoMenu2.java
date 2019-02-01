package tool_LoProgramming;
/* -*- Mode: Java; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This file incorporates work covered by the following license notice:
 *
 *   Licensed to the Apache Software Foundation (ASF) under one or more
 *   contributor license agreements. See the NOTICE file distributed
 *   with this work for additional information regarding copyright
 *   ownership. The ASF licenses this file to you under the Apache
 *   License, Version 2.0 (the "License"); you may not use this file
 *   except in compliance with the License. You may obtain a copy of
 *   the License at http://www.apache.org/licenses/LICENSE-2.0 .
 */

import com.sun.star.awt.MouseEvent;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.TextEvent;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XFocusListener;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XTextListener;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


public class UnoMenu2 extends UnoMenu {

public UnoMenu2(XComponentContext _xContext, XMultiComponentFactory _xMCF) {
    super(_xContext, _xMCF);
}

    public static void main(String args[]){
        UnoMenu2 oUnoMenu2 = null;
        try {
        XComponentContext xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
        if(xContext != null )
            System.out.println("Connected to a running office ...");
        XMultiComponentFactory xMCF = xContext.getServiceManager();
        oUnoMenu2 = new UnoMenu2(xContext, xMCF);
        oUnoMenu2.initialize( new String[] {"Height", "Moveable", "Name","PositionX","PositionY", "Step", "TabIndex","Title","Width"},
                                    new Object[] { Integer.valueOf(140), Boolean.TRUE, "Dialog1", Integer.valueOf(102),Integer.valueOf(41), Integer.valueOf(1), Short.valueOf((short) 0), "Menu-Dialog", Integer.valueOf(200)});

        Object oFTHeaderModel = oUnoMenu2.m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
        XMultiPropertySet xFTHeaderModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oFTHeaderModel);
        xFTHeaderModelMPSet.setPropertyValues(
            new String[] {"Height", "Label", "Name", "PositionX", "PositionY", "Width"},
            new Object[] { Integer.valueOf(8), "This code-sample demonstrates the creation of a popup-menu", "HeaderLabel", Integer.valueOf(6), Integer.valueOf(6), Integer.valueOf(200)});
        // add the model to the NameContainer of the dialog model
        oUnoMenu2.m_xDlgModelNameContainer.insertByName("Headerlabel", oFTHeaderModel);
        oUnoMenu2.addLabelForPopupMenu();
        oUnoMenu2.insertEditField();
        oUnoMenu2.executeDialog();
        }catch( Exception ex ) {
            ex.printStackTrace(System.err);
        }
        finally{
        //make sure always to dispose the component and free the memory!
        if (oUnoMenu2 != null) {
            if (oUnoMenu2.m_xComponent != null){
                oUnoMenu2.m_xComponent.dispose();
            }
        }
        System.exit( 0 );
    }}


    public void addLabelForPopupMenu(){
    try{
        String sName = "lblPopup";
        Object oFTModel = m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
        XMultiPropertySet xFTModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oFTModel);
        // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
        xFTModelMPSet.setPropertyValues(
            new String[] {"Height", "Label", "Name", "PositionX", "PositionY", "Width"},
            new Object[] { Integer.valueOf(8), "Right-click here", sName, Integer.valueOf(50), Integer.valueOf(50), Integer.valueOf(100)});
        // add the model to the NameContainer of the dialog model
        m_xDlgModelNameContainer.insertByName(sName, oFTModel);
        XWindow xWindow = UnoRuntime.queryInterface(XWindow.class, m_xDlgContainer.getControl(sName));
        xWindow.addMouseListener(this);
    }catch( Exception e ) {
        System.err.println( e + e.getMessage());
        e.printStackTrace();
    }}

    public void  insertEditField(){
        try{
            String sName = "TextField";
            Object oTFModel = m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlEditModel");
            XMultiPropertySet xTFModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oTFModel);

            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            xTFModelMPSet.setPropertyValues(
                    new String[] {"Height", "Name", "PositionX", "PositionY", "Text", "Width"},
                    new Object[] { Integer.valueOf(12), sName, Integer.valueOf(120), Integer.valueOf(80), "", Integer.valueOf(30)});

            m_xDlgModelNameContainer.insertByName(sName, oTFModel);
            XPropertySet xTFModelPSet = UnoRuntime.queryInterface(XPropertySet.class, oTFModel);

            XTextComponent xTextComponent = UnoRuntime.queryInterface(XTextComponent.class, m_xDlgContainer.getControl(sName));
            XWindow xTFWindow = UnoRuntime.queryInterface(XWindow.class, m_xDlgContainer.getControl(sName));
            xTextComponent.addTextListener(this);
        } catch (com.sun.star.uno.Exception ex) {
            ex.printStackTrace(System.err);
        }
    }

    @Override
    protected void closeDialog(){
        xDialog.endExecute();
    }

    @Override
    public void textChanged(TextEvent t) {
    	System.out.println("aaaa");
    }
    
    @Override
    public void mouseReleased(MouseEvent mouseEvent) {
    }

    @Override
    public void mousePressed(MouseEvent mouseEvent) {
        if (mouseEvent.PopupTrigger){
            Rectangle aPos = new Rectangle(mouseEvent.X, mouseEvent.Y, 0, 0);
            XControl xControl = UnoRuntime.queryInterface(XControl.class, mouseEvent.Source);
            getPopupMenu().execute( xControl.getPeer(), aPos, com.sun.star.awt.PopupMenuDirection.EXECUTE_DEFAULT);
        }
    }

    @Override
    public void mouseExited(MouseEvent mouseEvent) {
    }

    @Override
    public void mouseEntered(MouseEvent mouseEvent) {
    //	System.out.println("aaa");
    }

    @Override
    public void disposing(EventObject eventObject) {
    	System.out.println("aaa");
    	//Rectangle aPos = new Rectangle(mouseEvent.X, mouseEvent.Y, 0, 0);
        //XControl xControl = UnoRuntime.queryInterface(XControl.class, mouseEvent.Source);
        //getPopupMenu().execute( xControl.getPeer(), aPos, com.sun.star.awt.PopupMenuDirection.EXECUTE_DEFAULT);
    }
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
