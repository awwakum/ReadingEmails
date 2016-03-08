import java.util.*;

/**
 * Created by acvetkov on 07.03.2016.
 */
public class Machine implements Comparable<Machine> {
    private final String name;
    private TreeMap<String, String> dateValue = new TreeMap<>();

    public Machine(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public void setParams(String date, String value) {
        dateValue.put(date, value);
    }

    public void getAllParams() {
        for(String elem : dateValue.keySet()) {
            System.out.println(elem + " - " + dateValue.get(elem));
        }
    }

    public String getValueByKey(String key) {
        String value = "";
        for (String elem : dateValue.keySet()) {
            if (elem.equals(key)) value = dateValue.get(elem);
        }
        return value;
    }

    @Override
    public int compareTo(Machine m) {
        return name.compareTo(m.name);
    }

   /* public Map<String, String> getDateValue() {
        Set set = dateValue.entrySet();
        Iterator it = set.iterator();
    }*/
}
