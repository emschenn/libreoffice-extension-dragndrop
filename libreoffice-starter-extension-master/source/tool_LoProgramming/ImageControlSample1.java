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

public class ImageControlSample1 extends UnoDialogSample {
	/**
	 * Creates a new instance of ImageControlSample
	 */
	static ImageControlSample1 oImageControlSample;
	
	public ImageControlSample1(XComponentContext _xContext, XMultiComponentFactory _xMCF) {
		super(_xContext, _xMCF);
		super.createDialog(_xMCF);
	}

	// to start this script pass a parameter denoting the system path to a graphic
	// to be displayed
	public static void main(String args[]) {
		oImageControlSample = null;
		try {

			XComponentContext xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
			if (xContext != null) {
				System.out.println("Connected to a running office ...");
			}
			XMultiComponentFactory xMCF = xContext.getServiceManager();
			oImageControlSample = new ImageControlSample1(xContext, xMCF);
			
			oImageControlSample.initialize(
					new String[] { "Height", "Moveable", "Name", "PositionX", "PositionY", "Step", "TabIndex", "Title",
							"Width" },
					new Object[] { Integer.valueOf(300), Boolean.TRUE, "MyTestDialog", Integer.valueOf(102),
							Integer.valueOf(70), Integer.valueOf(0), Short.valueOf((short) 0), "My Gallery",
							Integer.valueOf(212) });
			// label
			Object oFTHeaderModel = oImageControlSample.m_xMSFDialogModel
					.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
			XMultiPropertySet xFTHeaderModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oFTHeaderModel);
			xFTHeaderModelMPSet.setPropertyValues(
					new String[] { "Height", "Label", "MultiLine", "Name", "PositionX", "PositionY", "Width" },
					new Object[] { Integer.valueOf(9), "Drag your image to the Text Field below:", Boolean.TRUE,
							"HeaderLabel", Integer.valueOf(6), Integer.valueOf(6), Integer.valueOf(195) });
			oImageControlSample.m_xDlgModelNameContainer.insertByName("Headerlabel", oFTHeaderModel);

			// image
			XPropertySet xICModelPropertySet = oImageControlSample.insertImageControl(6, 45, 90, 90);

			// insert text field
			XTextComponent text = oImageControlSample.insertTextField(oImageControlSample);
			text.addTextListener(new XTextListener() {
				@Override
				public void disposing(EventObject arg0) {
					// TODO Auto-generated method stub
				}

				@Override
				public void textChanged(TextEvent arg0) {
					String encodedURL = encode(text.getText().toString());
					System.out.println("Encoded URL: " + encodedURL);
					String decodedURL = decode(encodedURL);
					System.out.println("Decoded URL: " + decodedURL);
					String url = getImgStr(decodedURL);
					System.out.println("url: " + url);
					try {
						xICModelPropertySet.setPropertyValue("ImageURL", url);
					} catch (IllegalArgumentException | UnknownPropertyException | PropertyVetoException
							| WrappedTargetException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			});

			// file control
			XTextComponent file = oImageControlSample.insertFileControl(oImageControlSample, 6, 260, 160);
			//file.setText("D:\\研究擴充ㄉ冬\\gallery\\");
			file.addTextListener(new XTextListener() {
				@Override
				public void disposing(EventObject arg0) {
					// TODO Auto-generated method stub
				}

				@Override
				public void textChanged(TextEvent arg0) {
					// TODO Auto-generated method stub
					String encodedURL = encode(text.getText().toString());
					String decodedURL = decode(encodedURL);
					String url = getImgStr(decodedURL);
					System.out.println("url: " + url);
					System.out.println(file.getText());
				}
			});

			// gallery
			XPropertySet a[] = new XPropertySet[15];
			a[0] = oImageControlSample.insertImageControl(105, 45, 30, 30);
			a[1] = oImageControlSample.insertImageControl(140, 45, 30, 30);
			a[2] = oImageControlSample.insertImageControl(175, 45, 30, 30);
			a[3] = oImageControlSample.insertImageControl(105, 80, 30, 30);
			a[4] = oImageControlSample.insertImageControl(140, 80, 30, 30);
			a[5] = oImageControlSample.insertImageControl(175, 80, 30, 30);
			a[6] = oImageControlSample.insertImageControl(105, 115, 30, 30);
			a[7] = oImageControlSample.insertImageControl(140, 115, 30, 30);
			a[8] = oImageControlSample.insertImageControl(175, 115, 30, 30);
			a[9] = oImageControlSample.insertImageControl(105, 150, 30, 30);
			a[10] = oImageControlSample.insertImageControl(140, 150, 30, 30);
			a[11] = oImageControlSample.insertImageControl(175, 150, 30, 30);
			a[12] = oImageControlSample.insertImageControl(105, 185, 30, 30);
			a[13] = oImageControlSample.insertImageControl(140, 185, 30, 30);
			a[14] = oImageControlSample.insertImageControl(175, 185, 30, 30);
			
			
			//button
			XButton close = oImageControlSample.insertButton(oImageControlSample, 80, 280, 50, "~Close dialog",
					(short) PushButtonType.CANCEL_value);
			XButton gallery = oImageControlSample.insertButton(oImageControlSample, 6, 145, 50, "Add to ~Gallery",
					(short) PushButtonType.STANDARD_value);
			XButton slide = oImageControlSample.insertButton(oImageControlSample, 6, 165, 50, "Add to ~Slide",
					(short) PushButtonType.STANDARD_value);
			XButton check = oImageControlSample.insertButton(oImageControlSample, 170, 260, 30, "~Send",
					(short) PushButtonType.STANDARD_value);
			gallery.addActionListener(new XActionListener() {
				@Override
				public void disposing(EventObject arg0) {
					// TODO Auto-generated method stub
				}
				@Override
				public void actionPerformed(ActionEvent arg0) {
					URL url;
					try {
						String encodedURL = encode(text.getText().toString());
						String decodedURL = decode(encodedURL);
						String url1 = getImgStr(decodedURL);
						String name = getImgName(url1);
						url = new URL(url1);
						HttpURLConnection conn;
						conn = (HttpURLConnection) url.openConnection();
						conn.setRequestMethod("GET");
						conn.setConnectTimeout(5 * 1000);
						InputStream inStream = conn.getInputStream();
						byte[] data = readInputStream(inStream);
						File imageFile = new File(file.getText()+name);
						FileOutputStream outStream = new FileOutputStream(imageFile);
						outStream.write(data);
						outStream.close();
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			});
			check.addActionListener(new XActionListener() {
				@Override
				public void disposing(EventObject arg0) {
					// TODO Auto-generated method stub
				}
				@Override
				public void actionPerformed(ActionEvent arg0) {
					File f = new File(file.getText());
					File [] f1 = f.listFiles();
					for(int i=0;i<f.listFiles().length;i++) {
						System.out.println(f1[i]);
						XGraphic xGraphic = oImageControlSample.getGraphic(oImageControlSample.m_xMCF,f1[i].toString());
						try {
							a[i].setPropertyValue("Graphic", xGraphic);
						} catch (IllegalArgumentException | UnknownPropertyException | PropertyVetoException
								| WrappedTargetException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					}
					
				}
			});

			// label
			Object odModel = oImageControlSample.m_xMSFDialogModel
					.createInstance("com.sun.star.awt.UnoControlFixedTextModel");
			XMultiPropertySet xdModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, odModel);
			xdModelMPSet.setPropertyValues(
					new String[] { "Height", "Label", "MultiLine", "Name", "PositionX", "PositionY", "Width" },
					new Object[] { Integer.valueOf(9), "Destination of Gallery:", Boolean.TRUE, "DLabel",
							Integer.valueOf(6), Integer.valueOf(250), Integer.valueOf(210) });
			oImageControlSample.m_xDlgModelNameContainer.insertByName("Dlabel", odModel);

			oImageControlSample.createWindowPeer();

			// XGraphic xGraphic =
			// oImageControlSample.getGraphic(oImageControlSample.m_xMCF,
			// "https://i.imgur.com/qzY7BJ9.jpg");
			// xICModelPropertySet.setPropertyValue("Graphic", xGraphic);

			oImageControlSample.xDialog = UnoRuntime.queryInterface(XDialog.class,
					oImageControlSample.m_xDialogControl);
			oImageControlSample.executeDialog();

		} catch (Exception e) {
			System.err.println(e + e.getMessage());
			e.printStackTrace();
		} finally {
			// make sure always to dispose the component and free the memory!
			if (oImageControlSample != null) {
				if (oImageControlSample.m_xComponent != null) {
					oImageControlSample.m_xComponent.dispose();
				}
			}
		}
		System.exit(0);
	}

	public XTextComponent insertTextField(XTextListener _xTextListener) {
		XTextComponent xTextComponent = null;
		try {
			String sName = "TextField";
			Object oTFModel = m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlEditModel");
			XMultiPropertySet xTFModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oTFModel);
			// Set the properties at the model - keep in mind to pass the property names in
			// alphabetical order!
			xTFModelMPSet.setPropertyValues(
					new String[] { "Height", "Name", "PositionX", "PositionY", "Text", "Width" },
					new Object[] { Integer.valueOf(20), sName, Integer.valueOf(6), Integer.valueOf(18), "",
							Integer.valueOf(90) });
			m_xDlgModelNameContainer.insertByName(sName, oTFModel);
			// XPropertySet xTFModelPSet = UnoRuntime.queryInterface(XPropertySet.class, );
			XPropertySet xPropertySet = com.sun.star.uno.UnoRuntime.queryInterface(XPropertySet.class, oTFModel);
			// XTextComponent xTextComponent =
			// UnoRuntime.queryInterface(XTextComponent.class,
			// oImageControlSample.m_xDlgContainer.getControl(sName));
			XControl xFCControl = m_xDlgContainer.getControl(sName);
			xTextComponent = UnoRuntime.queryInterface(XTextComponent.class, xFCControl);
			XWindow xFCWindow = UnoRuntime.queryInterface(XWindow.class, xFCControl);
			xTextComponent.addTextListener(_xTextListener);
		} catch (com.sun.star.uno.Exception ex) {
			ex.printStackTrace(System.err);
		}
		return xTextComponent;
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
					new String[] { "Border", "Height", "Name", "PositionX", "PositionY", "ScaleImage", "Width", "Tabstop" },
					new Object[] { Short.valueOf((short) 1), Integer.valueOf(_nHeight), sName, Integer.valueOf(_nPosX),
							Integer.valueOf(_nPosY), Boolean.TRUE, Integer.valueOf(_nWidth), Boolean.TRUE });
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

	protected XTextComponent insertFileControl(XTextListener _xTextListener, int _nPosX, int _nPosY, int _nWidth) {
		XTextComponent xTextComponent = null;
		try {
			// create a unique name by means of an own implementation...
			String sName = createUniqueName(m_xDlgModelNameContainer, "FileControl");

			// retrieve the configured Work path...
			Object oPathSettings = thePathSettings.get(m_xContext);
			XPropertySet xPropertySet = com.sun.star.uno.UnoRuntime.queryInterface(XPropertySet.class, oPathSettings);
			String sWorkUrl = (String) xPropertySet.getPropertyValue("Work");

			// convert the Url to a system path that is "human readable"...
			Object oFCProvider = m_xMCF.createInstanceWithContext("com.sun.star.ucb.FileContentProvider", m_xContext);
			XFileIdentifierConverter xFileIdentifierConverter = UnoRuntime
					.queryInterface(XFileIdentifierConverter.class, oFCProvider);
			String sSystemWorkPath = xFileIdentifierConverter.getSystemPathFromFileURL(sWorkUrl);

			// create a controlmodel at the multiservicefactory of the dialog model...
			Object oFCModel = m_xMSFDialogModel.createInstance("com.sun.star.awt.UnoControlFileControlModel");

			// Set the properties at the model - keep in mind to pass the property names in
			// alphabetical order!
			XMultiPropertySet xFCModelMPSet = UnoRuntime.queryInterface(XMultiPropertySet.class, oFCModel);
			xFCModelMPSet.setPropertyValues(
					new String[] { "Height", "Name", "PositionX", "PositionY", "Text", "Width" },
					new Object[] { Integer.valueOf(14), sName, Integer.valueOf(_nPosX), Integer.valueOf(_nPosY),
							sSystemWorkPath, Integer.valueOf(_nWidth) });

			// The controlmodel is not really available until inserted to the Dialog
			// container
			m_xDlgModelNameContainer.insertByName(sName, oFCModel);
			XPropertySet xFCModelPSet = UnoRuntime.queryInterface(XPropertySet.class, oFCModel);

			// add a textlistener that is notified on each change of the controlvalue...
			XControl xFCControl = m_xDlgContainer.getControl(sName);
			xTextComponent = UnoRuntime.queryInterface(XTextComponent.class, xFCControl);
			XWindow xFCWindow = UnoRuntime.queryInterface(XWindow.class, xFCControl);
			xTextComponent.addTextListener(_xTextListener);
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
		return xTextComponent;
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
	public static String getImgName(String url) {
		String img = url.substring(url.lastIndexOf("/"));
		System.out.println(img);
		return img;
	}
	public static byte[] readInputStream(InputStream inStream) throws Exception {
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		byte[] buffer = new byte[1024];
		int len = 0;
		while ((len = inStream.read(buffer)) != -1) {
			outStream.write(buffer, 0, len);
		}
		inStream.close();
		return outStream.toByteArray();
	}

	
	
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
