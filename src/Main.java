import com.opencsv.CSVWriter;
import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.FolderSchema;
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

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.SortedSet;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    private static Pattern pattern =  Pattern.compile("([\\d]{0,5},[\\d]{0,2})");

    // allow Autodiscover to follow the redirection
    private static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
        public boolean autodiscoverRedirectionUrlValidationCallback(
                String redirectionUrl) {
            return redirectionUrl.toLowerCase().startsWith("https://");
        }
    }

    /* format date */
    private static String editDate(Date date) {
        SimpleDateFormat simpleDate = new SimpleDateFormat("dd.MM.YYYY");
        return simpleDate.format(date);
    }

    /* write to .csv file */
    private static void writeCSVFile(SortedSet<String> dateSet, Machine[] machinesArray) throws Exception {

        CSVWriter writer = new CSVWriter(new FileWriter("test.csv"), ';', CSVWriter.NO_QUOTE_CHARACTER);
        String[] record = "date,p14,p31,p32".split(",");
        writer.writeNext(record);

        for (String date : dateSet) {
            String str = date + "#";

            for (Machine machine : machinesArray) {
                /*System.out.println("Send getValueByKey with: " + date);*/
                str += machine.getValueByKey(date);
                str += "#";
            }

            record = str.split("#");
            writer.writeNext(record);
        }

        writer.close();
    }

    /* read the last date line in .csv file */
    /*public static String readCSVFile(String file) throws IOException {
        CSVReader reader = new CSVReader(new FileReader(file), ';');
        String[] nextLine;
        String lastDate = "";
        while ((nextLine = reader.readNext()) != null) {
            if (reader.readNext() == null) {
                lastDate = nextLine[0];
                break;
            }
        }
        return lastDate;
    }*/

    /* read properties file */
    private static Properties readPropertiesFile(String fileName) throws FileNotFoundException {
        FileInputStream fis;
        Properties prop;

        fis = new FileInputStream(fileName);
        prop = new Properties();

        try {
            prop.load(fis);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return prop;
    }

    public static void main(String[] args) throws Exception {

        /* run from console */
        /*String pass = "";
        try {
            pass = new String(System.console().readPassword("Password: "));
        } catch (NullPointerException e) {
            System.err.println("No console found");
            System.exit(1);
        }*/

        Properties prop = readPropertiesFile("config.properties");

        /* login part */
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        ExchangeCredentials credentials = new WebCredentials(prop.getProperty("user"), prop.getProperty("pass"));
        service.setCredentials(credentials);
        service.autodiscoverUrl(prop.getProperty("url"), new RedirectionUrlCallback());

        /* find needed folder in the inbox folder */
        SearchFilter.IsEqualTo filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, prop.getProperty("folderName"));
        FindFoldersResults findFoldersResults = service.findFolders(WellKnownFolderName.Inbox,
                filter, new FolderView(Integer.MAX_VALUE));
        String folderId = findFoldersResults.getFolders().get(0).getId().toString();

        /* bind to folder */
        Folder testFolder = Folder.bind(service, new FolderId(folderId));

        /* all messages in folder, date ascending sorted */
        ItemView view = new ItemView(Integer.MAX_VALUE);
        view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

        Machine[] machinesArray = {new Machine("p14"), new Machine("p31"), new Machine("p32")};
        SortedSet<String> dateSet = new TreeSet<>();


        FindItemsResults<Item> findResults = service.findItems(testFolder.getId(), view);
        for (Item item : findResults.getItems()) {
            EmailMessage message = EmailMessage.bind(service, new ItemId(item.getId().toString()));
            PropertySet itemPropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
            itemPropertySet.setRequestedBodyType(BodyType.Text);

            /* get values from message body */
            String machineName = "";
            try {
                machineName = item.getSubject().substring(32);
            } catch (IndexOutOfBoundsException e) {
                System.out.println(e);
            }

            /* we need only machines p14 p31 p32 */
            if(!machineName.equals("p14") && !machineName.equals("p31") && !machineName.equals("p32")) continue;

            String messageBody = message.getBody().toString();
            String date = editDate(item.getDateTimeReceived());
            /* add each date to dateSet */
            dateSet.add(date);
            String value;

            int beginIndex = messageBody.indexOf("Per Second");
            int endIndex = messageBody.indexOf("Remaining");

            if (beginIndex != -1) {
                Matcher matcher = pattern.matcher( messageBody.substring(beginIndex, endIndex));
                if(matcher.find()){
                    value = matcher.group();
                }
                else {
                    value=null;
                }
            } else continue;

            /* add params to each machine */
            switch (machineName) {
                case "p14" : machinesArray[0].setParams(date, value);
                    break;
                case "p31" : machinesArray[1].setParams(date, value);
                    break;
                case "p32" : machinesArray[2].setParams(date, value);
                    break;
            }
        } /* end for */

        /*System.out.println("\nMachines in ascending order: ");
        for (Machine machine : machinesArray) {
            System.out.println(machine.getName());
            machine.getAllParams();
        }*/

        writeCSVFile(dateSet, machinesArray);
    }
}