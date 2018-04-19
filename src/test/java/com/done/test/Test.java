package com.done.test;

import com.done.exception.CustomException;
import com.done.utils.JxlsUtils;

import java.util.*;

/**
 *
 */
public class Test {

    public static void main(String[] args) {
        List<Map<String,Object>> list = new ArrayList<Map<String, Object>>();
        Map<String,Object> map = new HashMap<String,Object>();
        Map<String,Object> map1 = new HashMap<String,Object>();
        for (int i = 0; i < 4; i++) {
            for (int j = 0; j < 8; j++) {
                map.put("key"+j,"value"+j);
            }
            list.add(map);
        }
        for (int m = 0; m < 13; m++) {
            map1.put("key"+m,"value"+m);
        }
        map1.put("nowdate",new Date());
        map1.put("phone","18513572398");
        System.out.println(list);
        System.out.println(map1);

        try {
            JxlsUtils.exportExcelToPath("E:\\custom\\custom.xls",list,map1,"E:\\custom\\custom-test01.xls");
        } catch (CustomException e) {
            e.printStackTrace();
        }
    }
}
