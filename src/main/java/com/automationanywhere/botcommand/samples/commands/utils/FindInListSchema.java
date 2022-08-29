package com.automationanywhere.botcommand.samples.commands.utils;

import com.automationanywhere.botcommand.data.model.Schema;

import java.util.ArrayList;
import java.util.List;

public class FindInListSchema {

    public List<Schema> schemas;

    //public FindInListSchema(List<Schema> objectList) { this.schemas = objectList; }

    public FindInListSchema(List<?> objectList) {
        if(objectList.size() >0) {
            if(objectList.get(0) instanceof String ) {
                List<String> stringList = this.cast(objectList);
                List<Schema> listSchema = new ArrayList();
                for (String sc : stringList) {
                    listSchema.add(new Schema(sc));
                }
                this.schemas = listSchema;
            }else{
                this.schemas = this.cast(objectList);
            }
        }
    }

    public Boolean exists(String value) {
        for(Schema lsc : this.schemas){
            if(lsc.getName().equals(value)){
                return true;
            }
        }
        return false;
    }
    public int indexSchema(String value) {
        int i = 0;
        for(Schema lsc : this.schemas){
            if(lsc.getName().equals(value)){
                return i;
            }
            i++;
        }
        return -1;
    }
    public List<Integer> indexSchema(List<String> value) {
        List<Integer> lista = new ArrayList();
        for(String v: value){
            lista.add(this.indexSchema(v));
//            int i = 0;
//            for(Schema lsc : this.schemas) {
//                if (lsc.getName().equals(v)) {
//                    lista.add(i);
//                }
//                i++;
//            }
        }
        return lista;
    }
    public List<String> shemaNames() {
        List<String> names = new ArrayList();


        for(Schema lsc : this.schemas){
            names.add(lsc.getName());
        }
        return names;
    }
    public void addSchema(Integer idx,String SchemaName) {
        Schema novo = new Schema(SchemaName);
        this.schemas.add(idx,novo);
    }
    public void addSchema(String SchemaName) {
        Schema novo = new Schema(SchemaName);
        this.schemas.add(novo);
    }
    public void alterSchema(Integer idx,String SchemaName) {
        Schema novo = new Schema(SchemaName);
        this.schemas.set(idx,novo);
    }
    public static <T> List<T> cast(List list) {
        return list;
    }
}
