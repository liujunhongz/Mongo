package tv.lycam.excel;

import tv.lycam.excel.utils.MongoUtils;
import tv.lycam.excel.utils.cglib.ExcelUtils;

import java.io.IOException;
import java.util.List;
import java.util.Scanner;

/**
 * @Author: 诸葛不亮
 * @Description: MongoDB数据库导出，Excel读取
 */
public class Main {


    public static void main(String[] args) throws IOException {
        if (args != null && args.length == 2) {
            String arg = args[0];
            switch (arg) {
                case "--import":
                    main_import(args[1]);
                    break;
                case "--export":
                    main_export();
                    break;
            }
            return;
        }
        main_export();
    }

    public static void main_export() throws IOException {
        Scanner scanner = new Scanner(System.in);
        System.out.println("此程序要完成的功能：Mongo数据到Excel，请完成下面5个步骤");
        System.out.println("(1)请输入mongo数据库IP地址,默认值：127.0.0.1,回车使用默认值");
        String ip = scanner.nextLine();
        if ((ip == null) || (ip.trim().isEmpty())) {
            ip = "127.0.0.1";
        }
        System.out.println(">>>>>>>>>你的输入：" + ip);
        System.out.println("(2)请输入mongo数据库端口,默认值：27017,回车使用默认值");
        String po = scanner.nextLine();
        int port = 27017;
        if ((po != null) && (!po.trim().isEmpty())) {
            port = Integer.valueOf(po);
        }
        System.out.println(">>>>>>>>>你的输入：" + port);
        System.out.println("(3)请输入要连接的mongo数据库,默认值：sbkt-dev,回车使用默认值");
        String database = scanner.nextLine();
        if ((database == null) || (database.trim().isEmpty())) {
            database = "sbkt-dev";
        }
        System.out.println(">>>>>>>>>你的输入：" + database);
        System.out.println("(4)请输入要查询的Collection,默认值：user,回车使用默认值");
        String collection = scanner.nextLine();
        if ((collection == null) || (collection.trim().isEmpty())) {
            collection = "user";
        }
        System.out.println(">>>>>>>>>你的输入：" + collection);
        MongoUtils.init(ip, port, database, collection);
//        System.out.println("(7)请输入公众号文件存储路径,默认值：./visit.txt,回车使用默认值");
//        String idsPath = scanner.nextLine();
//        if ((idsPath == null) || (idsPath.trim().isEmpty())) {
//            idsPath = "./visit.txt";
//        }
//        System.out.println(">>>>>>>>>你的输入：" + idsPath);
        List<String> jsons = MongoUtils.getAllJsonString();
        System.out.println("统计到的数量为：\n" + jsons.size());
        System.out.println("(5)请输入excel表格保存位置,默认值：./result.xls,回车使用默认值");
        String path = scanner.nextLine();
        if ((path == null) || (path.trim().isEmpty())) {
            path = "./result.xls";
        }
        System.out.println(">>>>>>>>>你的输入：" + path);
        ExcelUtils.exportAsExcel(jsons, path, collection);
        System.out.println("表格导出完成");
        MongoUtils.destory();
    }

    public static void main_import(String xls) throws IOException {
        System.out.println("读取xls文件：");
        List<Object> objs = ExcelUtils.importExcel(xls);
        for (Object obj : objs) {
            System.out.println(obj);
        }
    }

}
