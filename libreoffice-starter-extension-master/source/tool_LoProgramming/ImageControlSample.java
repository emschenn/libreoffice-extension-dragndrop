package tool_LoProgramming;

import java.awt.event.TextListener;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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

import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.TextEvent;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XTextListener;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameContainer;
import com.sun.star.datatransfer.dnd.DropTargetDragEnterEvent;
import com.sun.star.datatransfer.dnd.DropTargetDragEvent;
import com.sun.star.datatransfer.dnd.DropTargetDropEvent;
import com.sun.star.datatransfer.dnd.DropTargetEvent;
import com.sun.star.datatransfer.dnd.XDropTarget;
import com.sun.star.datatransfer.dnd.XDropTargetListener;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

public class ImageControlSample extends UnoDialogSample implements XDropTargetListener {
	/**
	 * Creates a new instance of ImageControlSample
	 */

    
	public ImageControlSample(XComponentContext _xContext, XMultiComponentFactory _xMCF) {
		super(_xContext, _xMCF);
		super.createDialog(_xMCF);
	}

	// to start this script pass a parameter denoting the system path to a graphic
	// to be displayed
	public static void main(String args[]) {
		//final ImageControlSample oImageControlSample = null;
		try {
			XComponentContext xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
			if (xContext != null) {
				System.out.println("Connected to a running office ...");
			}
			XMultiComponentFactory xMCF = xContext.getServiceManager();
			ImageControlSample oImageControlSample = new ImageControlSample(xContext, xMCF);
			oImageControlSample.initialize(
					new String[] { "Height", "Moveable", "Name", "PositionX", "PositionY", "Step", "TabIndex", "Title",
							"Width" },
					new Object[] { Integer.valueOf(100), Boolean.TRUE, "MyTestDialog", Integer.valueOf(102),
							Integer.valueOf(41), Integer.valueOf(0), Short.valueOf((short) 0), "OpenOffice",
							Integer.valueOf(230) });
			
			Object oFTHeaderModel = oImageControlSample.m_xMSFDialogModel
					.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
			XMultiPropertySet xFTHeaderModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oFTHeaderModel);
			xFTHeaderModelMPSet.setPropertyValues(
					new String[] { "Height", "Label", "MultiLine", "Name", "PositionX", "PositionY", "Width" },
					new Object[] { Integer.valueOf(16),
							"This code-sample demonstrates how to create an ImageControlSample within a dialog",
							Boolean.TRUE, "HeaderLabel", Integer.valueOf(6), Integer.valueOf(6),
							Integer.valueOf(210) });
			// add the model to the NameContainer of the dialog model
			oImageControlSample.m_xDlgModelNameContainer.insertByName("Headerlabel", oFTHeaderModel);

			
			// insert image control-----------------------------------------------------------
			XPropertySet xICModelPropertySet = oImageControlSample.insertImageControl(68, 30, 32, 90);
			
			// insert edit field-----------------------------------------------------------
			try {
				String sName = "TextField";
				Object oTFModel = m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlEditModel");
				XMultiPropertySet xTFModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oTFModel);

				// Set the properties at the model - keep in mind to pass the property names in
				// alphabetical order!
				xTFModelMPSet.setPropertyValues(
						new String[] { "Height", "Name", "PositionX", "PositionY", "Text", "Width" },
						new Object[] { Integer.valueOf(12), sName, Integer.valueOf(120), Integer.valueOf(80), "",
								Integer.valueOf(30) });

				oImageControlSample.m_xDlgModelNameContainer.insertByName(sName, oTFModel);
				XPropertySet xTFModelPSet = UnoRuntime.queryInterface(XPropertySet.class, oTFModel);

				XTextComponent xTextComponent = UnoRuntime.queryInterface(XTextComponent.class,
						oImageControlSample.m_xDlgContainer.getControl(sName));
				xTextComponent.addTextListener(new XTextListener() {
					@Override
					public void disposing(EventObject arg0) {
						// TODO Auto-generated method stub
					}
					@Override
					public void textChanged(TextEvent arg0) {
						String encodedURL = encode(xTextComponent.getText().toString());
						System.out.println("Encoded URL: " + encodedURL);
						String decodedURL = decode(encodedURL);
						System.out.println("Decoded URL: " + decodedURL);
						String url = getImgStr(decodedURL);
						System.out.println("url: " + url);
						//XGraphic xGraphic = oImageControlSample.getGraphic(oImageControlSample.m_xMCF, url);
						try {
							xICModelPropertySet.setPropertyValue("GraphicURL", url);
						} catch (IllegalArgumentException | UnknownPropertyException | PropertyVetoException
								| WrappedTargetException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					}
				});
			} catch (com.sun.star.uno.Exception ex) {
				ex.printStackTrace(System.err);
			}
			
			
			
			oImageControlSample.insertButton(oImageControlSample, 90, 75, 50, "~Close dialog",
					(short) PushButtonType.OK_value);
			oImageControlSample.createWindowPeer();
			// note: due to issue i76718 ("Setting graphic at a controlmodel required dialog
			// peer") the graphic of the image control
			// may not be set before the peer of the dialog has been created.

			// XGraphic xGraphic =
			// oImageControlSample.getGraphic(oImageControlSample.m_xMCF, args[0]);
			// xICModelPropertySet.setPropertyValue("Graphic", xGraphic);
			oImageControlSample.xDialog = UnoRuntime.queryInterface(XDialog.class,
					oImageControlSample.m_xDialogControl);
			oImageControlSample.executeDialog();

		} catch (Exception e) {
			System.err.println(e + e.getMessage());
			e.printStackTrace();
		} finally {
			// make sure always to dispose the component and free the memory!
			
		}
		System.exit(0);
	}

	
	
    public XPropertySet insertImageControl(int _nPosX, int _nPosY, int _nHeight, int _nWidth){
        XPropertySet xICModelPropertySet = null;
        try{
            // create a unique name by means of an own implementation...
            String sName = createUniqueName(m_xDlgModelNameContainer, "ImageControl");
            // convert the system path to the image to a FileUrl

            // create a controlmodel at the multiservicefactory of the dialog model...
            Object oICModel = m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlImageControlModel");
            XMultiPropertySet xICModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oICModel);
            xICModelPropertySet =UnoRuntime.queryInterface(XPropertySet.class, oICModel);
            // Set the properties at the model - keep in mind to pass the property names in alphabetical order!
            // The image is not scaled
            xICModelMPSet.setPropertyValues(
                    new String[] {"Border", "Height", "Name", "PositionX", "PositionY", "ScaleImage", "Width"},
                    new Object[] { Short.valueOf((short) 1), Integer.valueOf(_nHeight), sName, Integer.valueOf(_nPosX), Integer.valueOf(_nPosY), Boolean.FALSE, Integer.valueOf(_nWidth)});

            // The controlmodel is not really available until inserted to the Dialog container
            m_xDlgModelNameContainer.insertByName(sName, oICModel);
        }catch (com.sun.star.uno.Exception ex){
            /* perform individual exception handling here.
             * Possible exception types are:
             * com.sun.star.lang.IllegalArgumentException,
             * com.sun.star.lang.WrappedTargetException,
             * com.sun.star.container.ElementExistException,
             * com.sun.star.beans.PropertyVetoException,
             * com.sun.star.beans.UnknownPropertyException,
             * com.sun.star.uno.Exception
             */
            ex.printStackTrace(System.err);
        }
        return xICModelPropertySet;
    }
	
	
	
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
	// 百分比編碼函數
	public static String encode(String url) {
		try {
			String encodeURL = URLEncoder.encode(url, "UTF-8");
			return encodeURL;
		} catch (UnsupportedEncodingException e) {
			return "Error: " + e.getMessage();
		}
	}
	

	// creates a UNO graphic object that can be used to be assigned
	// to the property "Graphic" of a controlmodel
	public XGraphic getGraphic(XMultiComponentFactory _xMCF, String _sImageSystemPath) {
		XGraphic xGraphic = null;
		try {
			java.io.File oFile = new java.io.File(_sImageSystemPath);
			Object oFCProvider = _xMCF.createInstanceWithContext("com.sun.star.ucb.FileContentProvider",
					this.m_xContext);
			XFileIdentifierConverter xFileIdentifierConverter = UnoRuntime
					.queryInterface(XFileIdentifierConverter.class, oFCProvider);
			String sImageUrl = xFileIdentifierConverter.getFileURLFromSystemPath(_sImageSystemPath,
					oFile.getAbsolutePath());

			// create a GraphicProvider at the global service manager...
			Object oGraphicProvider = m_xMCF.createInstanceWithContext("com.sun.star.graphic.GraphicProvider",
					m_xContext);
			XGraphicProvider xGraphicProvider = UnoRuntime.queryInterface(XGraphicProvider.class, oGraphicProvider);
			// create the graphic object
			PropertyValue[] aPropertyValues = new PropertyValue[1];
			PropertyValue aPropertyValue = new PropertyValue();
			aPropertyValue.Name = "URL";
			aPropertyValue.Value = sImageUrl;
			aPropertyValues[0] = aPropertyValue;
			xGraphic = xGraphicProvider.queryGraphic(aPropertyValues);
			return xGraphic;
		} catch (com.sun.star.uno.Exception ex) {
			throw new java.lang.RuntimeException("cannot happen...", ex);
		}
	}


}

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
