package playground;

import java.lang.reflect.Field;
import java.util.*;


public class Example1 {
    public static void main(String[] args) {

        Map<Integer, String> map = new HashMap();

        for (int i = 0; i < 11 ;i++){
            map.put(i,"\t Value: "+i);
        }
        map.entrySet()
                .stream()
                .sorted(Map.Entry.comparingByValue())
                .forEach(System.out::println);


    }
    public static List<String> getFieldNames(Field[] fields){
        List<String> fieldNames = new ArrayList<>();
        for (Field field: fields){
            fieldNames.add(field.getName());
        }

        return fieldNames;
    }

}
