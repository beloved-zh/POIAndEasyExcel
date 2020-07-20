package com.zh.listen;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.fastjson.JSON;
import com.zh.pojo.User;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author Beloved
 * @date 2020/7/17 22:50
 * 创建实体类的读取
 */
public class UserListener extends AnalysisEventListener<User> {

    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 2;
    List<User> list = new ArrayList<User>();


    /**
     * 这个每一条数据解析都会来调用
     * @param user
     * @param analysisContext
     */
    public void invoke(User user, AnalysisContext analysisContext) {
        System.out.println("解析到一条数据:" + JSON.toJSONString(user));
        list.add(user);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (list.size() >= BATCH_COUNT) {
            //这里可以调用存储方法，存到数据库
            saveData();
            // 存储完成清理 list
            list.clear();
        }
    }

    /**
     * 这里会一行行的返回头
     *
     * @param headMap
     * @param context
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        System.out.println("解析到一条头数据:" + JSON.toJSONString(headMap));
    }

    /**
     * 所有数据解析完成了 都会来调用
     * @param analysisContext
     */
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
         saveData();
        System.out.println("所有数据解析完成！");
    }

    private void saveData() {
        System.out.println("开始存储数据库！");
        System.out.println(list);
    }


}
