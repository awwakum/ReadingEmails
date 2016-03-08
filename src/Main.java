/**
 * Created by acvetkov on 07.03.2016.
 */
import com.opencsv.CSVWriter;
import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

import java.io.Console;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Array;
import java.text.SimpleDateFormat;
import java.util.*;

public class Main {

    // allow Autodiscover to follow the redirection
    static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
        public boolean autodiscoverRedirectionUrlValidationCallback(
                String redirectionUrl) {
            return redirectionUrl.toLowerCase().startsWith("https://");
        }
    }

    // format date
    public static String editDate(Date date) {
        SimpleDateFormat simpleDate = new SimpleDateFormat("dd.MM.YYYY");
        return simpleDate.format(date);
    }

    public static void writeCSVFile(SortedSet dateSet, Machine[] machinesArray) throws Exception {

        CSVWriter writer = new CSVWriter(new FileWriter("test.csv"), ';', CSVWriter.NO_QUOTE_CHARACTER);
        String[] record = "date,p14,p31,p32".split(",");
        writer.writeNext(record);

        Iterator<String> it = dateSet.iterator();
        while (it.hasNext()) {

            String str = it.next() + "#";

            for (Machine machine : machinesArray) {
                str += machine.getValueByKey(it.next());
                str += "#";
            }

            record = str.split("#");
            writer.writeNext(record);
        }

        writer.close();
    }

    public static void main(String[] args) throws Exception {

        Scanner in = new Scanner(System.in);
        System.out.print("Password: ");
        String pass = in.nextLine();

        /* run from console */
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

        // find all child folders of the inbox folder
        FindFoldersResults findFoldersResults = service.findFolders(WellKnownFolderName.Inbox, new FolderView(Integer.MAX_VALUE));
        System.out.println("\nChild folders of the inbox folder: ");
        for (Folder folder : findFoldersResults.getFolders()) {
            System.out.println("Count = " + folder.getChildFolderCount());
            System.out.println("Name  = " + folder.getDisplayName());
            System.out.println("ID    = " + folder.getId());
        }

        Folder testFolder = Folder.bind(service, new FolderId("AAMkADljN2IwZDM2LWVhNWUtNDgzNi05YmE0LTZiOTQzMDhjYjA4ZgAuAAAAAAA4FlhziwSnQ51C62KpVWnkAQB5cYD3J63WTZwpPyMN8yhDAAAEqpRrAAA="));
        System.out.println("\nEmails in " + testFolder.getDisplayName());

        ItemView view = new ItemView(Integer.MAX_VALUE);
        view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

        //SortedSet<Machine> machinesSet = new TreeSet<>();
        Machine[] machinesArray = {new Machine("p14"), new Machine("p31"), new Machine("p32")};
        SortedSet<String> dateSet = new TreeSet<>();


        FindItemsResults<Item> findResults = service.findItems(testFolder.getId(), view);
        for (Item item : findResults.getItems()) {
            EmailMessage message = EmailMessage.bind(service, new ItemId(item.getId().toString()));
            PropertySet itemPropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
            itemPropertySet.setRequestedBodyType(BodyType.Text);

            /* get from message body */
            String machineName = item.getSubject().substring(32);
            /* we need only machines p14 p31 p32 */
            if(!machineName.equals("p14") && !machineName.equals("p31") && !machineName.equals("p32")) continue;
            String messageBody = message.getBody().toString();
            String date = editDate(item.getDateTimeReceived());
            String value = "";

            System.out.print(date + "\t" + machineName + "\t");
            int position = messageBody.indexOf("Per Second");
            if (position != -1) {
                value = messageBody.substring(position + 20, position + 25).replace("<", "");
                System.out.println(value);
            }
            else System.out.println("");

            /* add params to each machine */
            switch (machineName) {
                case "p14" : machinesArray[0].setParams(date, value);
                    break;
                case "p31" : machinesArray[1].setParams(date, value);
                    break;
                case "p32" : machinesArray[2].setParams(date, value);
                    break;
            }

            /*Machine machine = new Machine(machineName);
            machinesSet.add(machine);
            machine.setParams(date, value);*/

            /* add each date to dateSet */
            dateSet.add(date);
        } // end for

        System.out.println("Machines in ascending order: ");
        /*Iterator it = machinesSet.iterator();
        while(it.hasNext()) {
            Machine machine = (Machine) it.next();
            System.out.println(machine.getName());
            machine.getParams();
            System.out.println("=========================");
        }*/
        for (Machine machine : machinesArray) {
            System.out.println(machine.getName());
            machine.getAllParams();
        }

        writeCSVFile(dateSet, machinesArray);
    }
}