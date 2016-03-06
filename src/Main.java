import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import java.io.Console;
import java.util.Arrays;
import java.util.Scanner;


public class Main {

    static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
        public boolean autodiscoverRedirectionUrlValidationCallback(
                String redirectionUrl) {
            return redirectionUrl.toLowerCase().startsWith("https://");
        }
    }

    public static void main(String[] args) throws Exception {

        Scanner in = new Scanner(System.in);
        System.out.print("Password: ");
        String pass = in.nextLine();
        /*String pass = "";
        try {
            pass = new String(System.console().readPassword("Password: "));
        } catch (NullPointerException e) {
            System.err.println("No console found");
            System.exit(1);
        }*/

        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        ExchangeCredentials credentials = new WebCredentials("acvetkov", pass);
        service.setCredentials(credentials);
        //URI uri = URI.create ("https://access.freenet-group.de");
        //service.setUrl(uri);
        service.autodiscoverUrl("Alexander.Cvetkov@md.de", new RedirectionUrlCallback());

        Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);

        ItemView view = new ItemView(3);

        FindItemsResults<Item> findResults = service.findItems(inbox.getId(), view);
        for (Item item : findResults.getItems()) {
            System.out.println("id  : " + item.getId());
            System.out.println("sub : " + item.getSubject());
        }

        // find all child folders of the inbox folder
        FindFoldersResults findFoldersResults = service.findFolders(WellKnownFolderName.Inbox, new FolderView(Integer.MAX_VALUE));
        for (Folder folder : findFoldersResults.getFolders()) {
            System.out.println("Count = " + folder.getChildFolderCount());
            System.out.println("Name  = " + folder.getDisplayName());
            System.out.println("ID    = " + folder.getId());
        }

        //folder.delete(DeleteMode.HardDelete);

        EmailMessage message = EmailMessage.bind(service, new ItemId("AAMkADljN2IwZDM2LWVhNWUtNDgzNi05YmE0LTZiOTQzMDhjYjA4ZgBGAAAAAAA4FlhziwSnQ51C62KpVWnkBwB5cYD3J63WTZwpPyMN8yhDAAAA1TOBAAB5cYD3J63WTZwpPyMN8yhDAAAEqnOOAAA="));

        PropertySet itemPropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
        itemPropertySet.setRequestedBodyType(BodyType.Text);
        String body = message.getBody().toString();
        System.out.println(body);
    }
}