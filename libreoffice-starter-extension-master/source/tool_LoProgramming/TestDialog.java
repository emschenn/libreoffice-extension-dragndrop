package tool_LoProgramming;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.ProtocolException;
import java.net.URL;
import java.net.URLDecoder;
import java.net.URLEncoder;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.PushButtonType;
import com.sun.star.awt.TextEvent;
import com.sun.star.awt.XActionListener;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XTextListener;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XMultiPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.ucb.XFileIdentifierConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.thePathSettings;

public class TestDialog extends UnoDialogSample {
	/**
	 * Creates a new instance of ImageControlSample
	 */
	static TestDialog oTestDialog;

	public TestDialog(XComponentContext _xContext, XMultiComponentFactory _xMCF) {
		super(_xContext, _xMCF);
		super.createDialog(_xMCF);
	}

	// to start this script pass a parameter denoting the system path to a graphic
	// to be displayed
	public static void execute() {
		oTestDialog = null;
		try {
			XComponentContext xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
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
					new Object[] { Integer.valueOf(9), "the image is free to use", Boolean.FALSE,
							"HeaderLabel", Integer.valueOf(6), Integer.valueOf(27), Integer.valueOf(66) });
			oTestDialog.m_xDlgModelNameContainer.insertByName("Headerlabel", oFTHeaderModel);
			
			XPropertySet xICModelPropertySet = oTestDialog.insertImageControl(30, 6, 20, 20);
			xICModelPropertySet.setPropertyValue("ImageURL","http://icons.iconarchive.com/icons/paomedia/small-n-flat/1024/sign-check-icon.png" );
			// button
			XButton close = oTestDialog.insertButton(oTestDialog, 16, 40, 50, "~Ok",
					(short) PushButtonType.OK_value);
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
		System.exit(0);
	}
	public XPropertySet insertImageControl(int _nPosX, int _nPosY, int _nHeight, int _nWidth) {
		XPropertySet xICModelPropertySet = null;
		try {
			// create a unique name by means of an own implementation...
			String sName = createUniqueName(m_xDlgModelNameContainer, "ImageControl");
			// convert the system path to the image to a FileUrl
			// create a controlmodel at the multiservicefactory of the dialog model...
			Object oICModel = m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlImageControlModel");
			XMultiPropertySet xICModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oICModel);
			xICModelPropertySet = UnoRuntime.queryInterface(XPropertySet.class, oICModel);
			// Set the properties at the model - keep in mind to pass the property names in
			// alphabetical order!
			// The image is not scaled
			xICModelMPSet.setPropertyValues(
					new String[] { "Border", "Height", "Name", "PositionX", "PositionY", "ScaleImage", "Width"},
					new Object[] { Short.valueOf((short) 1), Integer.valueOf(_nHeight), sName, Integer.valueOf(_nPosX),
							Integer.valueOf(_nPosY), Boolean.TRUE, Integer.valueOf(_nWidth)});
			// The controlmodel is not really available until inserted to the Dialog
			// container
			
			m_xDlgModelNameContainer.insertByName(sName, oICModel);
		} catch (com.sun.star.uno.Exception ex) {
			/*
			 * perform individual exception handling here. Possible exception types are:
			 * com.sun.star.lang.IllegalArgumentException,
			 * com.sun.star.lang.WrappedTargetException,
			 * com.sun.star.container.ElementExistException,
			 * com.sun.star.beans.PropertyVetoException,
			 * com.sun.star.beans.UnknownPropertyException, com.sun.star.uno.Exception
			 */
			ex.printStackTrace(System.err);
		}
		return xICModelPropertySet;
	}
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
