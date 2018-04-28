package com.livedrof.poi.excel;

import com.livedrof.poi.data.User;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

public class ExporterTest {
    @Test
    public void testUserExporter() throws IOException {
        Exporter exporter = Exporter.getInstance();
        DataSheet<User> userDataSheet = exporter.sheet("用户信息");
        userDataSheet.addHeader("导出用户信息")
                .addHeader("共有1个用户");

        userDataSheet.getDataTable()
                .addColumn("ID", "id")
                .nextColumn("姓名", "username")
                .nextColumn("手机号", "telephone");
        userDataSheet.setData(this.getData());
        Workbook workbook = exporter.toWorkbook();
        File filename = new File("hello.xlsx");
        workbook.write(new FileOutputStream(filename));
    }

    protected List<User> getData() {
        User user = new User();
        user.setId(1);
        user.setUsername("jacky");
        user.setTelephone("14210064865");
        List<User> result = new LinkedList<>();
        result.add(user);
        return result;

    }
}
