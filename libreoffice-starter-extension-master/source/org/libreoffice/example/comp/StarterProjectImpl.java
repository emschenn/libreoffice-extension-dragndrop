package org.libreoffice.example.comp;

import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import com.sun.star.awt.Toolkit;

import tool.PageHelper;
import tool.ShapeHelper;
import tool_LoProgramming.*;

import java.awt.Image;
//import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.dnd.DropTarget;
import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.ImageIcon;
import javax.swing.JComponent;
import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.TransferHandler;

import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.lib.uno.helper.Factory;

import org.libreoffice.example.dialog.ActionOneDialog;
import org.libreoffice.example.helper.DialogHelper;
import org.libreoffice.example.helper.DocumentHelper;

import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.Point;
import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.Size;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XDataTransferProviderAccess;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XReschedule;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.ElementExistException;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.datatransfer.XTransferable;
import com.sun.star.datatransfer.dnd.DropTargetDragEnterEvent;
import com.sun.star.datatransfer.dnd.DropTargetDragEvent;
import com.sun.star.datatransfer.dnd.DropTargetEvent;
import com.sun.star.datatransfer.dnd.XDropTarget;
import com.sun.star.datatransfer.dnd.XDropTargetDragContext;
import com.sun.star.datatransfer.dnd.XDropTargetListener;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XDrawView;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.io.IOException;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.style.LineSpacing;
import com.sun.star.style.LineSpacingMode;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.lib.uno.helper.WeakBase;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;

/* -*- Mode: Java; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*************************************************************************
 *
 *  The Contents of this file are made available subject to the terms of
 *  the BSD license.
 *
 *  Copyright 2000, 2010 Oracle and/or its affiliates.
 *  All rights reserved.
 *
 *  Redistribution and use in source and binary forms, with or without
 *  modification, are permitted provided that the following conditions
 *  are met:
 *  1. Redistributions of source code must retain the above copyright
 *     notice, this list of conditions and the following disclaimer.
 *  2. Redistributions in binary form must reproduce the above copyright
 *     notice, this list of conditions and the following disclaimer in the
 *     documentation and/or other materials provided with the distribution.
 *  3. Neither the name of Sun Microsystems, Inc. nor the names of its
 *     contributors may be used to endorse or promote products derived
 *     from this software without specific prior written permission.
 *
 *  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
 *  "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 *  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
 *  FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
 *  COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 *  INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 *  BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS
 *  OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 *  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR
 *  TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
 *  USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 *************************************************************************/

//package com.sun.star.comp.sdk.examples;

import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.XActionListener;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlModel;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.comp.loader.FactoryHelper;
import com.sun.star.container.XNameContainer;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.task.XJobExecutor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.uno.Exception;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */

public final class StarterProjectImpl extends WeakBase
		implements com.sun.star.lang.XServiceInfo, com.sun.star.task.XJobExecutor {

	private static final String _buttonName = "Button1";
	private static final String _cancelButtonName = "CancelButton";
	private static final String _labelName = "Label1";
	private static final String _labelPrefix = "Number of button clicks: ";

	private Clipboard clipboard;

	private XComponentContext _xComponentContext;

	private ActionOneDialog actionOneDialog;

	private final XComponentContext m_xContext;

	private XMultiComponentFactory xMCF = null;
	private Object desktop = null;
	private XComponent xDrawDoc = null;
	private XTextDocument mxDoc = null;
	private XMultiServiceFactory mxDocFactory = null;
	private XDesktop xDesktop = null;

	private List ll;

	private static final String m_implementationName = StarterProjectImpl.class.getName();
	private static final String[] m_serviceNames = { "org.libreoffice.example.StarterProject" };

	public StarterProjectImpl(XComponentContext context) {
		m_xContext = context;

	};

	public static XSingleComponentFactory __getComponentFactory(String sImplementationName) {
		XSingleComponentFactory xFactory = null;

		if (sImplementationName.equals(m_implementationName))
			xFactory = Factory.createComponentFactory(StarterProjectImpl.class, m_serviceNames);
		return xFactory;
	}

	public static boolean __writeRegistryServiceInfo(XRegistryKey xRegistryKey) {
		return Factory.writeRegistryServiceInfo(m_implementationName, m_serviceNames, xRegistryKey);
	}

	// com.sun.star.lang.XServiceInfo:
	public String getImplementationName() {
		return m_implementationName;
	}

	public boolean supportsService(String sService) {
		int len = m_serviceNames.length;

		for (int i = 0; i < len; i++) {
			if (sService.equals(m_serviceNames[i]))
				return true;
		}
		return false;
	}

	public String[] getSupportedServiceNames() {
		return m_serviceNames;
	}

	// com.sun.star.task.XJobExecutor:
	public void trigger(String action) {
		switch (action) {
		case "actionOne":
			System.out.println("1");
			String pass = "the image is free to use";
			String error = "Unknown authorization";
			String passImg = "http://icons.iconarchive.com/icons/paomedia/small-n-flat/1024/sign-check-icon.png";
			String errorImg = "https://openclipart.org/image/2400px/svg_to_png/274092/warningSmall.png";

			xMCF = m_xContext.getServiceManager();
			if (xMCF == null) {
				System.err.println("ERROR: xMCF is null");
			}
			try {
				desktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", m_xContext);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			if (desktop == null) {
				System.err.println("ERROR: desktop is null");
			}
			xDrawDoc = DocumentHelper.getCurrentComponent(m_xContext);

			XDrawPage xPage = null;
			XShapes xShapes;
			XPropertySet xShapePropSet;
			XTextTable xTable = null;

			/*
			 * while (PageHelper.getDrawPageCount(xDrawDoc) < 5) try {
			 * PageHelper.insertNewDrawPageByIndex(xDrawDoc, 0); } catch
			 * (java.lang.Exception e) { // TODO Auto-generated catch block
			 * e.printStackTrace(); }
			 */
			System.out.println("doc start");

			try {
				xPage = PageHelper.getDrawPageByIndex(xDrawDoc, 0);
			} catch (IndexOutOfBoundsException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WrappedTargetException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			try {
				xPage = PageHelper.getDrawPageByIndex(xDrawDoc, 0);
			} catch (IndexOutOfBoundsException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WrappedTargetException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			String text2 = Draw.getShapeText(xPage);
			;
			System.out.println(text2);

			XFrame xframe = DocumentHelper.getCurrentFrame(m_xContext);
			if (xframe == null) {
				System.err.println("text null");
			}
			XDesktop xDesktop = DocumentHelper.getCurrentDesktop(m_xContext);
			com.sun.star.frame.XController xController = xframe.getController();
			XModel xModel = (XModel) xController.getModel();
			mxDoc = (XTextDocument) UnoRuntime.queryInterface(XTextDocument.class,
					DocumentHelper.getCurrentComponent(m_xContext));
			if (mxDoc == null) {
				System.err.println("text null");
			}

			// --------------------------------------------------------------------------------
			XDrawView xDrawView = UnoRuntime.queryInterface(XDrawView.class, xController);
			XDrawPage xDrawPage = xDrawView.getCurrentPage();
			XIndexAccess xIndexAccess2 = UnoRuntime.queryInterface(XIndexAccess.class, xDrawPage);
			XShape xShape = null;
			try {
				for (int i = 0; i < xIndexAccess2.getCount(); ++i) {
					if (i == xIndexAccess2.getCount() - 1)
						xShape = UnoRuntime.queryInterface(XShape.class, xIndexAccess2.getByIndex(i));
				}
			} catch (IndexOutOfBoundsException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (WrappedTargetException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			XTextRange xTextRange2 = UnoRuntime.queryInterface(XTextRange.class, xShape);
			System.out.println(xShape.getShapeType().toString());
			XText xText = xTextRange2.getText();
			System.out.println(xText.getString());

			String encodedURL = encode(xText.getString());
			System.out.println("Encoded URL: " + encodedURL);
			String decodedURL = decode(encodedURL);
			System.out.println("Decoded URL: " + decodedURL);
			String url = getImgStr(decodedURL);
			System.out.println("url " + url);

			// -------------------------------------------------------------------------------------
			try {
				ImageIcon img = new ImageIcon(new URL(url));
			} catch (MalformedURLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			System.out.println("img start");
			XMultiServiceFactory docServiceFactory = (XMultiServiceFactory) UnoRuntime
					.queryInterface(XMultiServiceFactory.class, xDrawDoc);
			if (docServiceFactory == null) {
				System.err.println("docServiceFactory null");
			}

			// Creating graphic shape service
			Object graphicShape = null;
			try {
				graphicShape = docServiceFactory.createInstance("com.sun.star.drawing.GraphicObjectShape");
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			if (graphicShape == null) {
				System.err.println("graphicShape null");
			}

			// Customizing graphic shape position and size
			XShape shapeSettings = (XShape) UnoRuntime.queryInterface(XShape.class, graphicShape);

			if (shapeSettings == null) {
				System.err.println("shapeSettings null");
			}

			try {
				shapeSettings.setSize(new Size(4680, 5080));
			} catch (PropertyVetoException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			shapeSettings.setPosition(new Point(20, 20));

			// ----------------

			XTransferable tt = UnoRuntime.queryInterface(XTransferable.class, xDrawDoc);
			if (tt == null) {
				System.err.println("qq");
			} else {
				System.out.println(tt.toString());
			}

			// ----------------

			// Creating bitmap container service
			XNameContainer bitmapContainer = null;
			String name = null;
			try {
				bitmapContainer = UnoRuntime.queryInterface(XNameContainer.class,
						docServiceFactory.createInstance("com.sun.star.drawing.BitmapTable"));
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			// Inserting test image to the container

			try {
				System.out.println(url);
				name = getImgName(url);
				System.out.println("name " + name);
				bitmapContainer.insertByName(name, url);
			} catch (IllegalArgumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (ElementExistException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WrappedTargetException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			/*
			 * try { bitmapContainer.insertByName(url, importData(tt));
			 * System.err.println("qqq"); } catch (Exception e) { // TODO Auto-generated
			 * catch block e.printStackTrace(); }
			 */

			// Querying property interface for the graphic shape service
			XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, graphicShape);

			try {
				System.out.println("Test img url is" + bitmapContainer.getByName(name));
			} catch (NoSuchElementException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WrappedTargetException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			// Assign test image internal URL to the graphic shape property
			try {
				xPropSet.setPropertyValue("GraphicURL", bitmapContainer.getByName(name));
				xPage.add(shapeSettings);
				xPage.remove(xShape);
			} catch (IllegalArgumentException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (UnknownPropertyException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (PropertyVetoException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (WrappedTargetException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (NoSuchElementException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			// Getting text field interface
			XText text = xTextRange2.getText();
			// Getting text cursor
			XTextCursor xTextCursor = text.createTextCursor();
			// Convert graphic shape to the text content item
			XTextContent xTextContent = (XTextContent) UnoRuntime.queryInterface(XTextContent.class, graphicShape);

			break;
		case "actionTwo":
			System.out.println("2");
			// PresentationDemo.main(null);

			break;
		case "actionThree":
			System.out.println("3");
			// PresentationDemo.main(null);

			break;
		default:
			DialogHelper.showErrorMessage(m_xContext, null, "Unknown action: " + action);
			System.out.println("error");
		}

	}

	private void createDialog() throws com.sun.star.uno.Exception {

		// get the service manager from the component context
		// change m_xContext
		XMultiComponentFactory xMultiComponentFactory = m_xContext.getServiceManager();

		// create the dialog model and set the properties
		Object dialogModel = xMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel",
				m_xContext);
		XPropertySet xPSetDialog = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, dialogModel);
		xPSetDialog.setPropertyValue("PositionX", Integer.valueOf(100));
		xPSetDialog.setPropertyValue("PositionY", Integer.valueOf(100));
		xPSetDialog.setPropertyValue("Width", Integer.valueOf(150));
		xPSetDialog.setPropertyValue("Height", Integer.valueOf(100));
		xPSetDialog.setPropertyValue("Title", "Runtime Dialog Demo");

		// get the service manager from the dialog model
		XMultiServiceFactory xMultiServiceFactory = (XMultiServiceFactory) UnoRuntime
				.queryInterface(XMultiServiceFactory.class, dialogModel);

		// create the button model and set the properties
		Object buttonModel = xMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");
		XPropertySet xPSetButton = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, buttonModel);
		xPSetButton.setPropertyValue("PositionX", Integer.valueOf(20));
		xPSetButton.setPropertyValue("PositionY", Integer.valueOf(70));
		xPSetButton.setPropertyValue("Width", Integer.valueOf(50));
		xPSetButton.setPropertyValue("Height", Integer.valueOf(14));
		xPSetButton.setPropertyValue("Name", _buttonName);
		xPSetButton.setPropertyValue("TabIndex", Short.valueOf((short) 0));
		xPSetButton.setPropertyValue("Label", "Click Me");

		// create the label model and set the properties
		Object labelModel = xMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
		XPropertySet xPSetLabel = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, labelModel);
		xPSetLabel.setPropertyValue("PositionX", Integer.valueOf(40));
		xPSetLabel.setPropertyValue("PositionY", Integer.valueOf(30));
		xPSetLabel.setPropertyValue("Width", Integer.valueOf(100));
		xPSetLabel.setPropertyValue("Height", Integer.valueOf(14));
		xPSetLabel.setPropertyValue("Name", _labelName);
		xPSetLabel.setPropertyValue("TabIndex", Short.valueOf((short) 1));
		xPSetLabel.setPropertyValue("Label", _labelPrefix);

		// create a Cancel button model and set the properties
		Object cancelButtonModel = xMultiServiceFactory.createInstance("com.sun.star.awt.UnoControlButtonModel");
		XPropertySet xPSetCancelButton = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class,
				cancelButtonModel);
		xPSetCancelButton.setPropertyValue("PositionX", Integer.valueOf(80));
		xPSetCancelButton.setPropertyValue("PositionY", Integer.valueOf(70));
		xPSetCancelButton.setPropertyValue("Width", Integer.valueOf(50));
		xPSetCancelButton.setPropertyValue("Height", Integer.valueOf(14));
		xPSetCancelButton.setPropertyValue("Name", _cancelButtonName);
		xPSetCancelButton.setPropertyValue("TabIndex", Short.valueOf((short) 2));
		xPSetCancelButton.setPropertyValue("PushButtonType", Short.valueOf((short) 2));
		xPSetCancelButton.setPropertyValue("Label", "Cancel");

		// insert the control models into the dialog model
		XNameContainer xNameCont = (XNameContainer) UnoRuntime.queryInterface(XNameContainer.class, dialogModel);
		xNameCont.insertByName(_buttonName, buttonModel);
		xNameCont.insertByName(_labelName, labelModel);
		xNameCont.insertByName(_cancelButtonName, cancelButtonModel);

		// create the dialog control and set the model
		Object dialog = xMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.UnoControlDialog",
				m_xContext);
		XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, dialog);
		XControlModel xControlModel = (XControlModel) UnoRuntime.queryInterface(XControlModel.class, dialogModel);
		xControl.setModel(xControlModel);

		// add an action listener to the button control
		XControlContainer xControlCont = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, dialog);
		Object objectButton = xControlCont.getControl("Button1");
		XButton xButton = (XButton) UnoRuntime.queryInterface(XButton.class, objectButton);
		xButton.addActionListener(new ActionListenerImpl(xControlCont));

		// create a peer
		Object toolkit = xMultiComponentFactory.createInstanceWithContext("com.sun.star.awt.ExtToolkit", m_xContext);
		XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, toolkit);
		XWindow xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, xControl);
		xWindow.setVisible(false);
		xControl.createPeer(xToolkit, null);

		// execute the dialog
		XDialog xDialog = (XDialog) UnoRuntime.queryInterface(XDialog.class, dialog);
		xDialog.execute();

		// dispose the dialog
		XComponent xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, dialog);
		xComponent.dispose();
	}

	public class ActionListenerImpl implements com.sun.star.awt.XActionListener {

		private int _nCounts = 0;

		private XControlContainer _xControlCont;

		public ActionListenerImpl(XControlContainer xControlCont) {
			_xControlCont = xControlCont;
		}

		// XEventListener
		public void disposing(EventObject eventObject) {
			_xControlCont = null;
		}

		// XActionListener
		public void actionPerformed(ActionEvent actionEvent) {

			// increase click counter
			_nCounts++;

			// set label text
			Object label = _xControlCont.getControl("Label1");
			XFixedText xLabel = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, label);
			xLabel.setText(_labelPrefix + _nCounts);
		}
	}

	public static com.sun.star.text.XTextDocument openWriter(com.sun.star.uno.XComponentContext xContext) {
		// define variables
		com.sun.star.frame.XComponentLoader xCLoader;
		com.sun.star.text.XTextDocument xDoc = null;
		com.sun.star.lang.XComponent xComp = null;

		try {
			// get the remote office service manager
			com.sun.star.lang.XMultiComponentFactory xMCF = xContext.getServiceManager();

			Object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);

			xCLoader = UnoRuntime.queryInterface(com.sun.star.frame.XComponentLoader.class, oDesktop);
			com.sun.star.beans.PropertyValue[] szEmptyArgs = new com.sun.star.beans.PropertyValue[0];
			String strDoc = "private:factory/swriter";
			xComp = xCLoader.loadComponentFromURL(strDoc, "_blank", 0, szEmptyArgs);
			xDoc = UnoRuntime.queryInterface(com.sun.star.text.XTextDocument.class, xComp);

		} catch (Exception e) {
			System.err.println(" Exception " + e);
			e.printStackTrace(System.err);
		}
		return xDoc;
	}

	public static void insertIntoCell(String CellName, String theText, com.sun.star.text.XTextTable xTTbl) {

		com.sun.star.text.XText xTableText = UnoRuntime.queryInterface(com.sun.star.text.XText.class,
				xTTbl.getCellByName(CellName));

		// create a cursor object
		com.sun.star.text.XTextCursor xTC = xTableText.createTextCursor();

		com.sun.star.beans.XPropertySet xTPS = UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, xTC);

		try {
			xTPS.setPropertyValue("CharColor", Integer.valueOf(16777215));
		} catch (Exception e) {
			System.err.println(" Exception " + e);
			e.printStackTrace(System.err);
		}

		// inserting some Text
		xTableText.setString(theText);

	}

	public static String getImgStr(String htmlStr) {
		Set<String> pics = new HashSet<>();
		String img = "";
		Pattern p_image = null;
		Matcher m_image = null;
		// String regEx_img = "<img.*src=(.*?)[^>]*?>"; // 圖片鏈接地址
		String regEx_img = "imgurl*=*.*imgrefurl";

		p_image = Pattern.compile(regEx_img, Pattern.CASE_INSENSITIVE);
		if (p_image == null) {
			System.err.println("p null");
		}
		m_image = p_image.matcher(htmlStr);
		if (m_image == null) {
			System.err.println("m null");
		}

		while (m_image.find()) {
			img = m_image.group();
			img = img.substring(7, img.length() - 10);
			System.out.println(img);
			pics.add(img);
		}
		/*
		 * while (m_image.find()) { // 得到<img />數據 img = m_image.group(); //
		 * 匹配<img>中的src數據 Matcher m = Pattern.compile ("\\s*=\\s*\"?(.*?)(\"|>|\\s+)"
		 * ).matcher(img); while (m.find()) { pics.add(m.group( 1 )); } }
		 */
		return img;
	}

	public static String getImgName(String htmlStr) {
		Set<String> pics = new HashSet<>();
		String img = "";
		Pattern p_image = null;
		Matcher m_image = null;
		// String regEx_img = "<img.*src=(.*?)[^>]*?>"; // 圖片鏈接地址
		String regEx_img = "/.*.(jp(e)?g)?(png)?";

		p_image = Pattern.compile(regEx_img, Pattern.CASE_INSENSITIVE);
		if (p_image == null) {
			System.err.println("p null");
		}
		m_image = p_image.matcher(htmlStr);
		if (m_image == null) {
			System.err.println("m null");
		}

		while (m_image.find()) {
			img = m_image.group();
			img = img.substring(7, img.length() - 10);
			System.out.println(img);
			pics.add(img);
		}
		/*
		 * while (m_image.find()) { // 得到<img />數據 img = m_image.group(); //
		 * 匹配<img>中的src數據 Matcher m = Pattern.compile ("\\s*=\\s*\"?(.*?)(\"|>|\\s+)"
		 * ).matcher(img); while (m.find()) { pics.add(m.group( 1 )); } }
		 */
		return img;
	}

	// 百分比解碼函數
	public static String decode(String url) {
		try {
			String prevURL = "";
			String decodeURL = url;
			while (!prevURL.equals(decodeURL)) {
				prevURL = decodeURL;
				decodeURL = URLDecoder.decode(decodeURL, "UTF-8");
			}
			return decodeURL;
		} catch (UnsupportedEncodingException e) {
			return "Error: " + e.getMessage();
		}
	}

	// 百分比編碼函數
	public static String encode(String url) {
		try {
			String encodeURL = URLEncoder.encode(url, "UTF-8");
			return encodeURL;
		} catch (UnsupportedEncodingException e) {
			return "Error: " + e.getMessage();
		}
	}

	public static byte[] ImageToByteArray(String url) throws java.io.IOException {
		URL u = new URL(url);
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		InputStream is = null;
		byte[] byteChunk = new byte[4096];
		try {
			is = new BufferedInputStream(u.openStream());
			int n;
			while ((n = is.read(byteChunk)) > 0) {
				baos.write(byteChunk, 0, n);
			}
		} finally {
			if (is != null) {
				is.close();
			}
		}

		return baos.toByteArray();
	}

	public ImageIcon importData(XTransferable tt) {
		for (com.sun.star.datatransfer.DataFlavor flavor : tt.getTransferDataFlavors()) {
			System.out.println(flavor);

			System.out.println("xxx  " + flavor.MimeType);
			if (DataFlavor.imageFlavor.equals(flavor)) {
				Object o = null;
				try {
					o = tt.getTransferData((com.sun.star.datatransfer.DataFlavor) flavor);
				} catch (com.sun.star.datatransfer.UnsupportedFlavorException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				if (o instanceof Image) {
					return (new ImageIcon((Image) o));
				}
			}
			if (DataFlavor.javaFileListFlavor.equals(flavor)) {
				Object o = null;
				try {
					o = (Transferable) tt.getTransferData(flavor);
				} catch (com.sun.star.datatransfer.UnsupportedFlavorException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				if (o instanceof List) {
					List list = (List) o;
					for (Object f : list) {
						if (f instanceof File) {
							File file = (File) f;
							System.out.println(file);
							if (!file.getName().endsWith(".bmp")) {
								return (new ImageIcon(file.getAbsolutePath()));
							}
						}
					}
				}
			}
		}
		System.out.println("gg");
		return null;
	}

	private void TestDialog(String s1, String s2) throws com.sun.star.uno.Exception {
		TestDialog oTestDialog = null;
		try {
			XComponentContext xContext = m_xContext;
			if (xContext != null) {
				System.out.println("Connected to a running office ...");
			}
			XMultiComponentFactory xMCF = xContext.getServiceManager();
			oTestDialog = new TestDialog(xContext, xMCF);
			oTestDialog.initialize(
					new String[] { "Height", "Moveable", "Name", "PositionX", "PositionY", "Step", "TabIndex", "Title",
							"Width" },
					new Object[] { Integer.valueOf(60), Boolean.TRUE, "MyTestDialog", Integer.valueOf(400),
							Integer.valueOf(80), Integer.valueOf(0), Short.valueOf((short) 0), "CC0 Test",
							Integer.valueOf(80) });
			// label
			Object oFTHeaderModel = oTestDialog.m_xMSFDialogModel
					.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
			XMultiPropertySet xFTHeaderModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oFTHeaderModel);
			xFTHeaderModelMPSet.setPropertyValues(
					new String[] { "Height", "Label", "MultiLine", "Name", "PositionX", "PositionY", "Width" },
					new Object[] { Integer.valueOf(9), s1, Boolean.FALSE, "HeaderLabel", Integer.valueOf(6),
							Integer.valueOf(27), Integer.valueOf(66) });
			oTestDialog.m_xDlgModelNameContainer.insertByName("Headerlabel", oFTHeaderModel);

			XPropertySet xICModelPropertySet = oTestDialog.insertImageControl(30, 6, 20, 20);
			xICModelPropertySet.setPropertyValue("ImageURL", s2);
			// button
			XButton close = oTestDialog.insertButton(oTestDialog, 16, 40, 50, "~Ok", (short) PushButtonType.OK_value);
			oTestDialog.createWindowPeer();
			oTestDialog.xDialog = UnoRuntime.queryInterface(XDialog.class, oTestDialog.m_xDialogControl);
			oTestDialog.executeDialog();

		} catch (Exception e) {
			System.err.println(e + e.getMessage());
			e.printStackTrace();
		} finally {
			// make sure always to dispose the component and free the memory!
			if (oTestDialog != null) {
				if (oTestDialog.m_xComponent != null) {
					oTestDialog.m_xComponent.dispose();
				}
			}
		}
	}

}
