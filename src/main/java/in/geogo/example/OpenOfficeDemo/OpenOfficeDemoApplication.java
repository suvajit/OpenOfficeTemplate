package in.geogo.example.OpenOfficeDemo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.XComponentContext;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.container.XSet;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.uno.XInterface;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextFieldsSupplier;
import com.sun.star.text.XText;
import com.sun.star.text.XTextRange;
import com.sun.star.container.XNameAccess;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XBookmarksSupplier;
import com.sun.star.beans.XPropertySet;
import com.sun.star.beans.XPropertyState;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.util.XRefreshable;

import java.util.Hashtable;

import ooo.connector.BootstrapSocketConnector;


@SpringBootApplication
public class OpenOfficeDemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(OpenOfficeDemoApplication.class, args);
		try {
		    OpenOfficeDemoApplication.templateExample();
		} catch( Exception e) {
					e.printStackTrace(System.err);
		}


	}


	/** Load a document as template */
  protected static XComponent newDocComponentFromTemplate(String loadUrl) throws java.lang.Exception {

		 //XComponentContext xContext = null;

     // Set path to OOO installation
		 String oooExeFolder = "/Applications/OpenOffice.app/Contents/MacOS";

     // Get the UNO bootstrap
     //xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();
		 XComponentContext xContext = BootstrapSocketConnector.bootstrap(oooExeFolder);

		 // get the remote service manager
		 XMultiComponentFactory xMCF = xContext.getServiceManager();

		 // retrieve the Desktop object, we need its XComponentLoader
		 Object oDesktop = xMCF.createInstanceWithContext(
																 "com.sun.star.frame.Desktop", xContext);

		 // get the component laoder from the desktop to create a new text document
		 XComponentLoader xComponentLoader = (XComponentLoader) UnoRuntime.queryInterface(
																				 XComponentLoader.class,oDesktop);

		 // define load properties according to com.sun.star.document.MediaDescriptor
		 // the boolean property AsTemplate tells the office to create a new document
		 // from the given file
		 PropertyValue[] loadProps = new PropertyValue[1];
		 loadProps[0] = new PropertyValue();
		 loadProps[0].Name = "AsTemplate";
		 loadProps[0].Value = new Boolean(true);
		 // load
		 return xComponentLoader.loadComponentFromURL(loadUrl, "_blank", 0, loadProps);
 }

 /** Sample for use of templates
      This sample uses the file TextTemplateWithUserFields.odt from the Samples folder.
      The file contains a number of User text fields (Variables - User) and a bookmark
      which we use to fill in various values
   */
  protected static void templateExample() throws java.lang.Exception {
      // create a small hashtable that simulates a rowset with columns
      Hashtable recipient = new Hashtable();
      recipient.put("Company", "Manatee Books");
      recipient.put("Contact", "Rod Martin");
      recipient.put("ZIP", "34567");
      recipient.put("City", "Fort Lauderdale");
      recipient.put("State", "Florida");

      // load template with User fields and bookmark
      XComponent xTemplateComponent = newDocComponentFromTemplate(
      "file:///Users/suvajit/Works/OpenOfficeDemo/TextTemplateWithUserFields.odt");

      // get XTextFieldsSupplier and XBookmarksSupplier interfaces from document component
      XTextFieldsSupplier xTextFieldsSupplier = (XTextFieldsSupplier)UnoRuntime.queryInterface(
          XTextFieldsSupplier.class, xTemplateComponent);
      XBookmarksSupplier xBookmarksSupplier = (XBookmarksSupplier)UnoRuntime.queryInterface(
          XBookmarksSupplier.class, xTemplateComponent);

      // access the TextFields and the TextFieldMasters collections
      XNameAccess xNamedFieldMasters = xTextFieldsSupplier.getTextFieldMasters();
      XEnumerationAccess xEnumeratedFields = xTextFieldsSupplier.getTextFields();

      // iterate over hashtable and insert values into field masters
      java.util.Enumeration keys = recipient.keys();
      while (keys.hasMoreElements()) {
          // get column name
          String key = (String)keys.nextElement();

          // access corresponding field master
          Object fieldMaster = xNamedFieldMasters.getByName(
              "com.sun.star.text.fieldmaster.User." + key);

          // query the XPropertySet interface, we need to set the Content property
          XPropertySet xPropertySet = (XPropertySet)UnoRuntime.queryInterface(
              XPropertySet.class, fieldMaster);

          // insert the column value into field master
          xPropertySet.setPropertyValue("Content", recipient.get(key));
      }

      // afterwards we must refresh the textfields collection
      XRefreshable xRefreshable = (XRefreshable)UnoRuntime.queryInterface(
          XRefreshable.class, xEnumeratedFields);
      xRefreshable.refresh();

      // accessing the Bookmarks collection of the document
      XNameAccess xNamedBookmarks = xBookmarksSupplier.getBookmarks();

      // find the bookmark named "Subscription"
      Object bookmark = xNamedBookmarks.getByName("Subscription");

      // we need its XTextRange which is available from getAnchor(),
      // so query for XTextContent
      XTextContent xBookmarkContent = (XTextContent)UnoRuntime.queryInterface(
          XTextContent.class, bookmark);

      // get the anchor of the bookmark (its XTextRange)
      XTextRange xBookmarkRange = xBookmarkContent.getAnchor();

      // set string at the bookmark position
      xBookmarkRange.setString("subscription for the Manatee Journal");
  }

}
