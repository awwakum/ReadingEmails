import java.util.*;

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
            if (elem.equals(key)) {
                value = dateValue.get(elem);
                /*System.out.println("Elem: " + elem + ", Value: " + value);*/
            }
        }
        return value;
    }

    @Override
    public int compareTo(Machine m) {
        return name.compareTo(m.name);
    }
}
