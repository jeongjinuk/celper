package org.celper.core.reader;

import java.util.ArrayList;
import java.util.List;

public class CSVRecord {

    private final List<String> record = new ArrayList<>();

    public CSVRecord() {}

    public void addField(String field){
        record.add(field);
    }

    public List<String> getRecord() {
        return record;
    }
}
