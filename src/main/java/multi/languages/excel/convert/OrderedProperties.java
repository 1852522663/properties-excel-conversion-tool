package multi.languages.excel.convert;

import java.util.Collections;
import java.util.Enumeration;
import java.util.LinkedHashSet;
import java.util.Properties;
import java.util.Set;

/**
 Java 的 Properties 加载属性文件后是无法保证输出的顺序与文件中一致的，
 因为 Properties 是继承自 Hashtable 的， key/value 都是直接存在 Hashtable 中的，
 而 Hashtable 是不保证进出顺序的。 此处覆盖原来Properties 写新的 OrderedProperties
 */
public class OrderedProperties extends Properties {
    private final LinkedHashSet<Object> keys = new LinkedHashSet<Object>();

    public Enumeration<Object> keys() {
        return Collections.<Object> enumeration(keys);
    }


    public Object put(Object key, Object value) {
        keys.add(key);
        return super.put(key, value);
    }


     

}
