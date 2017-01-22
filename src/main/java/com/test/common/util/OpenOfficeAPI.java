package com.test.common.util;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.style.XStyle;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.*;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Type;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import ooo.connector.BootstrapSocketConnector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * 主要用openOffice 的API向加密的word文件中写入文字和在图片;
 *
 *
 */
public class OpenOfficeAPI {
    private XComponentContext mxRemoteContext;
    private XMultiComponentFactory mxRemoteServiceManager;
    private XTextCursor mxDocCursor;
    private XText mxDocText;
    private XTextDocument mxDoc;
    private XSentenceCursor xSentenceCursor;
    private Object desktop;
    private XComponent xEmptyWriterComponent;
    /**
     * open office的安装路径;
     */
    private String OPENOFFICE_PATH ="/usr/bin/";

    private static Logger logger = LoggerFactory.getLogger(OpenOfficeAPI.class);

    public OpenOfficeAPI(String openOfficePath)
    {
        this.OPENOFFICE_PATH=openOfficePath;
    }

    /**
     * get the remote service manager
     *
     * @return
     * @throws java.lang.Exception
     */
    private XMultiComponentFactory getRemoteServiceManager()
            throws java.lang.Exception {
        if (mxRemoteContext == null && mxRemoteServiceManager == null) {
            // get the remote office context
            mxRemoteContext = BootstrapSocketConnector
                    .bootstrap(OPENOFFICE_PATH);
            if (logger.isInfoEnabled()) {
                logger.info("Connected to a running office ...");
            }
            mxRemoteServiceManager = mxRemoteContext.getServiceManager();
            String available = (mxRemoteServiceManager != null ? "available"
                    : "not available");
            if (logger.isInfoEnabled()) {
                logger.info("remote ServiceManager is " + available);
            }
        }
        return mxRemoteServiceManager;
    }

    /**
     * 判断文件是否存在,并将文件路径格式化为:file:///D:/doc/myword.doc的形式;
     *
     * @param fileUrl
     * @return
     * @throws Exception
     */
    private String formatFileUrl(String fileUrl) throws Exception {
        try {
            StringBuffer sUrl = new StringBuffer("file:///");
            File sourceFile = new File(fileUrl);
            sUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
            return sUrl.toString();
        } catch (Exception e) {
            if (logger.isInfoEnabled()) {
                logger.info("formatFileUrl=" + e.toString() + "fileUrl="
                        + fileUrl);
            }
        }
        return "";
    }

    /**
     * get the interfaces to control the UNO
     *
     * @param fileUrl
     * @return
     * @throws java.lang.Exception
     */
    private XComponent newDocComponent(String fileUrl, String password)
            throws java.lang.Exception {
        mxRemoteServiceManager = this.getRemoteServiceManager();
        // get the Desktop service
        desktop = mxRemoteServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", mxRemoteContext);
        // retrieve the current component and access the controller
        XComponentLoader xComponentLoader = (XComponentLoader) UnoRuntime
                .queryInterface(XComponentLoader.class, desktop);
        // set the OpenOffice not open
        HashMap<String, Object> map = new HashMap<String, Object>();
        map.put("Hidden", Boolean.TRUE);
        if( !password.trim().isEmpty()){
            map.put("Password", password);
        }
        PropertyValue[] propertyValue = MakePropertyValue(map);
        fileUrl = formatFileUrl(fileUrl);
        return xComponentLoader.loadComponentFromURL(fileUrl, "_blank", 0,
                propertyValue);
    }

    /**
     * 保存文件
     *
     * @param xDoc
     * @param storeUrl
     * @throws java.lang.Exception
     */
    private void storeDocComponent(XComponent xDoc, String storeUrl,
                                   String password) throws java.lang.Exception {
        XStorable xStorable = (XStorable) UnoRuntime.queryInterface(
                XStorable.class, xDoc);
        HashMap<String, Object> map = new HashMap<String, Object>();
        map.put("FilterName", "MS Word 97");
        if((password.trim().isEmpty())){
            map.put("Password", password);
        }
        PropertyValue[] storeProps = MakePropertyValue(map);
        storeUrl = formatFileUrl(storeUrl);
        xStorable.storeAsURL(storeUrl, storeProps);
    }

    /**
     * 关闭组件;
     *
     * @throws Exception
     */
    private void close() throws Exception {
        com.sun.star.util.XCloseable xCloseable = (com.sun.star.util.XCloseable) UnoRuntime
                .queryInterface(com.sun.star.util.XCloseable.class, mxDoc);
        if (xCloseable != null) {
            xCloseable.close(false);
        } else {
            com.sun.star.lang.XComponent xComponent = (com.sun.star.lang.XComponent) UnoRuntime
                    .queryInterface(com.sun.star.lang.XComponent.class, mxDoc);
            if (null != xComponent)
                xComponent.dispose();
        }
    }

    /**
     * 设置属性值;
     *
     * @param map
     * @return
     */
    private final PropertyValue[] MakePropertyValue(HashMap<String, Object> map) {
        int total = 0;
        if (null != map && map.size() > 0) {
            total = map.size();
        }
        PropertyValue[] tempMakePropertyValue = new PropertyValue[total];
        Iterator iter = map.entrySet().iterator();
        int i = 0;
        while (iter.hasNext()) {
            Map.Entry entry = (Map.Entry) iter.next();
            String key = entry.getKey().toString();
            Object val = entry.getValue();
            tempMakePropertyValue[i] = new PropertyValue();
            tempMakePropertyValue[i].Name = key;
            tempMakePropertyValue[i].Value = val;
            i = i + 1;
        }
        return tempMakePropertyValue;
    }

    /**
     * 插入字符串到word文件中;
     * @param fileUrl:word文件物理路径
     * @param content:文件中插入的内容
     * @param password:打开word文件的密码
     * @throws java.lang.Exception
     */
    public void insertStrToWord(String fileUrl, String content, String password)
            throws Exception {
        try {
            xEmptyWriterComponent = newDocComponent(fileUrl, password);
            mxDoc = (XTextDocument) UnoRuntime.queryInterface(
                    XTextDocument.class, xEmptyWriterComponent);
            mxDocText = mxDoc.getText();
            mxDocCursor = mxDocText.createTextCursor();
            xSentenceCursor = (XSentenceCursor) UnoRuntime.queryInterface(
                    XSentenceCursor.class, mxDocCursor);
            mxDocText.insertString(xSentenceCursor, content, true);
            // 保存;
            storeDocComponent(xEmptyWriterComponent, fileUrl, password);
        } catch (Exception e) {
            if (logger.isInfoEnabled()) {
                logger.info("插入字符串到word文件中" + e);
            }
        } finally {
            close();
        }

    }

    /**
     * 插入图片到word文件中;
     * @param fileUrl:word文件物理路径
     * @param imgUrl:文件中要插入的图片的物理路径
     * @param password:打开word文件的密码
     * @param width:插入的图片在word文件中的宽度
     * @param height:插入的图片在word文件中的高度
     * @throws java.lang.Exception
     */
    public void insertImgToWord(String fileUrl, String imgUrl, Integer width,
                                Integer height, String password) throws Exception {

        try {
            xEmptyWriterComponent = newDocComponent(fileUrl, password);
            mxDoc = (XTextDocument) UnoRuntime.queryInterface(
                    XTextDocument.class, xEmptyWriterComponent);
            mxDocText = mxDoc.getText();
            mxDocCursor = mxDocText.createTextCursor();
            xSentenceCursor = (XSentenceCursor) UnoRuntime.queryInterface(
                    XSentenceCursor.class, mxDocCursor);
            // 图片textcontent
            XMultiServiceFactory xMSFDoc = (XMultiServiceFactory) UnoRuntime
                    .queryInterface(XMultiServiceFactory.class,
                            xEmptyWriterComponent);
            Object oGraphic = xMSFDoc
                    .createInstance("com.sun.star.text.TextGraphicObject");
            XTextContent xTextContent = (XTextContent) UnoRuntime
                    .queryInterface(XTextContent.class, oGraphic);
            // 图片属性设置;
            XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, oGraphic);
            xPropSet.setPropertyValue("AnchorType",
                    TextContentAnchorType.AT_PARAGRAPH_value);
            imgUrl = formatFileUrl(imgUrl);
            xPropSet.setPropertyValue("GraphicURL", imgUrl);
            xPropSet.setPropertyValue("Width", new Integer(width * 10));
            xPropSet.setPropertyValue("Height", new Integer(height * 10));
            // 插入图片;
            mxDocText.insertTextContent(xSentenceCursor, xTextContent, true);
            // 保存;
            storeDocComponent(xEmptyWriterComponent, fileUrl, password);
        } catch (Exception e) {
            if (logger.isInfoEnabled()) {
                logger.info("插入图片到word文件中" + e);
            }
        } finally {
            close();
        }

    }

    /**
     * 插入图片到word页尾中;
     * @param  fileUrl:word文件绝对路径
     * @param  password:word文件密码,如果无密码则空;如果有密码,则新生成的newFileUrl,也用该密码加密
     * @param  imgUrl:插入图片文件绝对路径
     * @param  width:图片在word中展示的宽度
     * @param  height:图片在word中展示的高度
     * @param  newFileUrl:新生成文件的绝对路径
     * @throws java.lang.Exception
     * @return 返回1:生成成功;返回0:生成异常;
     */
    public int insertImgToFooter(String fileUrl,String password,String imgUrl, Integer width,
                                 Integer height, String newFileUrl) throws Exception {
        int rev=1;
        try {
            xEmptyWriterComponent = newDocComponent(fileUrl, password);
            mxDoc = (XTextDocument) UnoRuntime.queryInterface(
                    XTextDocument.class, xEmptyWriterComponent);
            XStyleFamiliesSupplier StyleFam = (XStyleFamiliesSupplier) UnoRuntime
                    .queryInterface(XStyleFamiliesSupplier.class, mxDoc);
            XNameAccess StyleFamNames = StyleFam.getStyleFamilies();
            XNameAccess PageStyles = (XNameAccess) AnyConverter.toObject(
                    new Type(XNameAccess.class),
                    StyleFamNames.getByName("PageStyles"));
            XStyle StdStyle = (XStyle) AnyConverter.toObject(new Type(
                    XStyle.class), PageStyles.getByName("Standard"));
            XPropertySet PropSet = (XPropertySet) UnoRuntime.queryInterface(
                    XPropertySet.class, StdStyle);
            PropSet.setPropertyValue("FooterIsOn", new Boolean(true));
            // Pegando o HeaderXText
            XText footerText = (com.sun.star.text.XText) UnoRuntime
                    .queryInterface(com.sun.star.text.XText.class,
                            PropSet.getPropertyValue("FooterTextLeft"));

            XMultiServiceFactory xMSFDoc = (XMultiServiceFactory) UnoRuntime
                    .queryInterface(XMultiServiceFactory.class, mxDoc);
            Object oGraphic = xMSFDoc
                    .createInstance("com.sun.star.text.TextGraphicObject");
            // Getting the cursor on the document
            XTextCursor xTextCursor = footerText.createTextCursor();
            XTextContent xTextContent = (XTextContent) UnoRuntime
                    .queryInterface(XTextContent.class, oGraphic);

            xTextCursor.gotoStart(false);
            // isReplace
            footerText.insertTextContent(xTextCursor, xTextContent, false);
            XPropertySet xPropSet = (com.sun.star.beans.XPropertySet) UnoRuntime
                    .queryInterface(XPropertySet.class, oGraphic);
            // Setting the anchor type
            xPropSet.setPropertyValue("AnchorType",
                    TextContentAnchorType.AT_PARAGRAPH);
            imgUrl = formatFileUrl(imgUrl);
            // Setting the graphic url
            xPropSet.setPropertyValue("GraphicURL", imgUrl);
            // Setting the horizontal position
            xPropSet.setPropertyValue("TextWrap", WrapTextMode.DYNAMIC);
            xPropSet.setPropertyValue("HoriOrient", HoriOrientation.LEFT);
            // Setting the width
            xPropSet.setPropertyValue("Width", width * 10);
            // Setting the height
            xPropSet.setPropertyValue("Height", height * 10);

            // 保存;
            storeDocComponent(xEmptyWriterComponent, newFileUrl, password);
        } catch (Exception e) {
            if (logger.isInfoEnabled()) {
                logger.info("插入图片到word页尾中" + e);
            }
            rev=0;
        } finally {
            close();
        }
        return rev;

    }

}