import java.util.*;

/**
 * Created by acvetkov on 07.03.2016.
 */
public class Machine {
    final private String name;
    private TreeMap<String, String> dateValue = new TreeMap<>();

    public Machine(String name, String date, String value) {
        this.name = name;
        dateValue.put(date, value);
    }

    public String getName() {
        return name;
    }

   /* public Map<String, String> getDateValue() {
        Set set = dateValue.entrySet();
        Iterator it = set.iterator();
    }*/
}
