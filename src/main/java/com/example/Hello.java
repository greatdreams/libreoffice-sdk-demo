package com.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.*;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.test.common.util.OpenOfficeAPI;
import ooo.connector.BootstrapSocketConnector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;

/** the main class of the program
 *
 *@author greatdreams
 */
public class Hello {
    private static Logger LOGGER = LoggerFactory.getLogger(Hello.class);
    /** the entry method of the program
     * @param args application arguments
     */
    public static void main(String[] args) throws Exception {
        LOGGER.info("This application is build with the sbt tool");
        XComponentContext xCtx = null;
        try {
            System.out.println(Bootstrap.class.getClassLoader().getParent().getClass());
            xCtx = Bootstrap.bootstrap();
            // String folder= "/usr/bin";
            // xCtx = BootstrapSocketConnector.bootstrap(folder);
            if(xCtx != null) {
                System.out.println("Connected to a running node...");
            }
        }catch (Exception e) {
            e.printStackTrace(System.out);
        }

        System.out.println("Opening an empty Writer document");

        XTextDocument myDoc = openWriter(xCtx);

        XText xText = myDoc.getText();
        XTextCursor xTCursor = xText.createTextCursor();
        xTCursor.gotoEnd(true);
        XSentenceCursor xSentenceCursor = (XSentenceCursor) UnoRuntime.queryInterface(
                XSentenceCursor.class, xTCursor);
        xText.insertString(xSentenceCursor, "This is my first document.\n", false);
        xText.insertString(xTCursor, "Now It's over\n", false);

        LOGGER.debug("Inserting a text table...");
        //getting MSF of the document
        XMultiServiceFactory xDocMSF = UnoRuntime.queryInterface(XMultiServiceFactory.class, myDoc);
        //create instance of a text table
        XTextTable xTT = null;
        try {
            Object oInt = xDocMSF.createInstance("com.sun.star.text.TextTable");
            xTT = UnoRuntime.queryInterface(XTextTable.class, oInt);
        }catch (Exception e) {
            e.printStackTrace(System.out);
        }
        //initialize the text table with 4 columns an 4 rows
        xTT.initialize(4,4);

        com.sun.star.beans.XPropertySet xTTRowPS = null;

        //insert the table
        try {
            xText.insertTextContent(xTCursor, xTT, false);
            // get first Row
            com.sun.star.container.XIndexAccess xTTRows = xTT.getRows();
            xTTRowPS = UnoRuntime.queryInterface(
                    com.sun.star.beans.XPropertySet.class, xTTRows.getByIndex(0));

        } catch (Exception e) {
            System.err.println("Couldn't insert the table " + e);
            e.printStackTrace(System.err);
        }

        // get the property set of the text table

        com.sun.star.beans.XPropertySet xTTPS = UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, xTT);

        // Change the BackColor
        try {
            xTTPS.setPropertyValue("BackTransparent", Boolean.FALSE);
            xTTPS.setPropertyValue("BackColor",Integer.valueOf(13421823));
            xTTRowPS.setPropertyValue("BackTransparent", Boolean.FALSE);
            xTTRowPS.setPropertyValue("BackColor",Integer.valueOf(6710932));

        } catch (Exception e) {
            System.err.println("Couldn't change the color " + e);
            e.printStackTrace(System.err);
        }

        // write Text in the Table headers
        System.out.println("Write text in the table headers");

        insertIntoCell("A1","FirstColumn", xTT);
        insertIntoCell("B1","SecondColumn", xTT) ;
        insertIntoCell("C1","ThirdColumn", xTT) ;
        insertIntoCell("D1","SUM", xTT) ;


        //Insert Something in the text table
        System.out.println("Insert something in the text table");

        (xTT.getCellByName("A2")).setValue(22.5);
        (xTT.getCellByName("B2")).setValue(5615.3);
        (xTT.getCellByName("C2")).setValue(-2315.7);
        (xTT.getCellByName("D2")).setFormula("sum <A2:C2>");

        (xTT.getCellByName("A3")).setValue(21.5);
        (xTT.getCellByName("B3")).setValue(615.3);
        (xTT.getCellByName("C3")).setValue(-315.7);
        (xTT.getCellByName("D3")).setFormula("sum <A3:C3>");

        (xTT.getCellByName("A4")).setValue(121.5);
        (xTT.getCellByName("B4")).setValue(-615.3);
        (xTT.getCellByName("C4")).setValue(415.7);
        (xTT.getCellByName("D4")).setFormula("sum <A4:C4>");


        //oooooooooooooooooooooooooooStep 5oooooooooooooooooooooooooooooooooooooooo
        // insert a colored text.
        // Get the propertySet of the cursor, change the CharColor and add a
        // shadow. Then insert the Text via InsertString at the cursor position.


        // get the property set of the cursor
        com.sun.star.beans.XPropertySet xTCPS = UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class,
                xTCursor);

        // Change the CharColor and add a Shadow
        try {
            xTCPS.setPropertyValue("CharColor",Integer.valueOf(255));
            xTCPS.setPropertyValue("CharShadowed", Boolean.TRUE);
        } catch (Exception e) {
            System.err.println("Couldn't change the color " + e);
            e.printStackTrace(System.err);
        }

        //create a paragraph break
        try {
            xText.insertControlCharacter(xTCursor,
                    com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false);

        } catch (Exception e) {
            System.err.println("Couldn't insert break "+ e);
            e.printStackTrace(System.err);
        }

        //inserting colored Text
        System.out.println("Inserting colored Text");

        xText.insertString(xTCursor, " This is a colored Text - blue with shadow\n",
                false );

        //create a paragraph break
        try {
            xText.insertControlCharacter(xTCursor,
                    com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false);

        } catch (Exception e) {
            System.err.println("Couldn't insert break "+ e);
            e.printStackTrace(System.err);
        }

        //oooooooooooooooooooooooooooStep 6oooooooooooooooooooooooooooooooooooooooo
        // insert a text frame.
        // create an instance of com.sun.star.text.TextFrame using the MSF of the
        // document. Change some properties an insert it.
        // Now get the text-Object of the frame an the corresponding cursor.
        // Insert some text via insertString.


        // Create a TextFrame
        com.sun.star.text.XTextFrame xTF = null;
        com.sun.star.drawing.XShape xTFS = null;

        try {
            Object oInt = xDocMSF.createInstance("com.sun.star.text.TextFrame");
            xTF = UnoRuntime.queryInterface(
                    com.sun.star.text.XTextFrame.class,oInt);
            xTFS = UnoRuntime.queryInterface(
                    com.sun.star.drawing.XShape.class,oInt);

            com.sun.star.awt.Size aSize = new com.sun.star.awt.Size();
            aSize.Height = 400;
            aSize.Width = 15000;

            xTFS.setSize(aSize);
        } catch (Exception e) {
            System.err.println("Couldn't create instance "+ e);
            e.printStackTrace(System.err);
        }

        // get the property set of the text frame
        com.sun.star.beans.XPropertySet xTFPS = UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, xTF);

        // Change the AnchorType
        try {
            xTFPS.setPropertyValue("AnchorType",
                    com.sun.star.text.TextContentAnchorType.AS_CHARACTER);
        } catch (Exception e) {
            System.err.println("Couldn't change the color " + e);
            e.printStackTrace(System.err);
        }

        //insert the frame
        System.out.println("Insert the text frame");

        try {
            xText.insertTextContent(xTCursor, xTF, false);
        } catch (Exception e) {
            System.err.println("Couldn't insert the frame " + e);
            e.printStackTrace(System.err);
        }

        //getting the text object of Frame
        com.sun.star.text.XText xTextF = xTF.getText();

        //create a cursor object
        com.sun.star.text.XTextCursor xTCF = xTextF.createTextCursor();

        //inserting some Text
        xTextF.insertString(xTCF,
                "The first line in the newly created text frame.", false);


        xTextF.insertString(xTCF,
                "\nWith this second line the height of the frame raises.", false);

        //insert a paragraph break
        try {
            xText.insertControlCharacter(xTCursor,
                    com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false );

        } catch (Exception e) {
            System.err.println("Couldn't insert break "+ e);
            e.printStackTrace(System.err);
        }

        // Change the CharColor and add a Shadow
        try {
            xTCPS.setPropertyValue("CharColor",Integer.valueOf(65536));
            xTCPS.setPropertyValue("CharShadowed", Boolean.FALSE);
        } catch (Exception e) {
            System.err.println("Couldn't change the color " + e);
            e.printStackTrace(System.err);
        }

        xText.insertString(xTCursor, " That's all for now !!", false );

        System.out.println("done");

        /*
        OpenOfficeAPI openOfficeAPI = new OpenOfficeAPI("/usr/bin");
        openOfficeAPI.insertStrToWord("/home/greatdreams/temp/temp.docx",
                "first document",
                "123456");
        */
    }

    public static XTextDocument openWriter(XComponentContext xCtx) {
        XComponentLoader xloader;
        XTextDocument xDoc = null;
        XComponent xComp = null;

        try {
            XMultiComponentFactory xMultiComponentFactory = xCtx.getServiceManager();
            System.out.println(xMultiComponentFactory.getAvailableServiceNames().length);
            java.lang.Object oDesktop = xMultiComponentFactory.createInstanceWithContext("com.sun.star.frame.Desktop",
                    xCtx);
            /*
            // get the Desktop service
            desktop = mxRemoteServiceManager.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", mxRemoteContext);
            */
            System.out.println("oDesktop:" + oDesktop == null?null: oDesktop.toString());
            xloader = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, oDesktop);
            System.out.println(xloader == null ? "null": xloader);
            PropertyValue [] szEmptyArgs = new PropertyValue[0];
            File file = new File("/home/greatdreams/temp/temp.docx");
            xComp = xloader.loadComponentFromURL(formatFileUrl(file.getAbsolutePath()), "_blank",0, szEmptyArgs);
            xDoc = UnoRuntime.queryInterface(XTextDocument.class, xComp);
        }catch (Exception e) {
            e.printStackTrace(System.out);
        }
        return xDoc;
    }
    public static String formatFileUrl(String fileUrl) throws Exception {
        try {
            StringBuffer sUrl = new StringBuffer("file:///");
            File sourceFile = new File(fileUrl);
            sUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
            return sUrl.toString();
        } catch (Exception e) {
        }
        return "";
    }

    public static void insertIntoCell(String CellName, String theText,
                                      com.sun.star.text.XTextTable xTTbl) {

        com.sun.star.text.XText xTableText = UnoRuntime.queryInterface(com.sun.star.text.XText.class,
                xTTbl.getCellByName(CellName));

        //create a cursor object
        com.sun.star.text.XTextCursor xTC = xTableText.createTextCursor();

        com.sun.star.beans.XPropertySet xTPS = UnoRuntime.queryInterface(com.sun.star.beans.XPropertySet.class, xTC);

        try {
            xTPS.setPropertyValue("CharColor",Integer.valueOf(16777215));
        } catch (Exception e) {
            System.err.println(" Exception " + e);
            e.printStackTrace(System.err);
        }

        //inserting some Text
        xTableText.setString( theText );

    }
}
