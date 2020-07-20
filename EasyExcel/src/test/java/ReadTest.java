import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.zh.listen.ExcelListener;
import com.zh.listen.UserListener;
import com.zh.pojo.User;
import org.junit.Test;

import java.io.File;

/**
 * @author Beloved
 * @date 2020/7/17 23:00
 */
public class ReadTest {

    @Test
    public void test01(){

        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 写法1：
        String fileName = "E:\\ideaProject\\POIAndEasyExcel\\User.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, User.class, new UserListener()).sheet().doRead();

    }

    @Test
    public void test02(){
        // 写法2：
        String fileName = "E:\\ideaProject\\POIAndEasyExcel\\User.xlsx";
        ExcelReader excelReader = null;
        try {
            excelReader = EasyExcel.read(fileName, User.class, new UserListener()).build();
            ReadSheet readSheet = EasyExcel.readSheet(0).build();
            excelReader.read(readSheet);
        } finally {
            if (excelReader != null) {
                // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
                excelReader.finish();
            }
        }
    }

    /**
     *
     * 不创建实体类
     */
    @Test
    public void test03(){
        String fileName = "E:\\ideaProject\\POIAndEasyExcel\\User.xlsx";
        // 这里 只要，然后读取第一个sheet 同步读取会自动finish
        EasyExcel.read(fileName, new ExcelListener()).sheet().doRead();
    }
}
