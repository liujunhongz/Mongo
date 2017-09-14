package tv.lycam.excel.utils;

import com.alibaba.fastjson.JSON;
import com.mongodb.MongoClient;
import com.mongodb.MongoException;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import com.mongodb.client.model.Filters;
import org.bson.BsonUndefined;
import org.bson.Document;
import org.bson.conversions.Bson;
import tv.lycam.excel.utils.cglib.CglibUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class MongoUtils {
    private static MongoClient mg = null;
    private static MongoDatabase db;
    private static MongoCollection<Document> visit;

    public static void init(String ip, int port, String database, String collection) {
        try {
            //连接到MongoDB服务 如果是远程连接可以替换“localhost”为服务器所在IP地址
            //ServerAddress()两个参数分别为 服务器地址 和 端口
//            ServerAddress serverAddress = new ServerAddress(ip,port);
//            List<ServerAddress> addrs = new ArrayList<ServerAddress>();
//            addrs.add(serverAddress);

            //MongoCredential.createScramSha1Credential()三个参数分别为 用户名 数据库名称 密码
//            MongoCredential credential = MongoCredential.createScramSha1Credential("username", "databaseName", "password".toCharArray());
//            List<MongoCredential> credentials = new ArrayList<MongoCredential>();
//            credentials.add(credential);
//            mg = new MongoClient(addrs);
            mg = new MongoClient(ip, port);
        } catch (MongoException e) {
            e.printStackTrace();
        }
        db = mg.getDatabase(database);
        visit = db.getCollection(collection);
    }

    public static void destory() {
        if (mg != null) {
            mg.close();
        }
        mg = null;
        db = null;
        visit = null;
        System.gc();
    }

    public static List<String> getAllJsonString() throws IOException {
        String[] keys = CglibUtils.getPropertyMap().keySet().toArray(new String[]{});
        List<Bson> bsons = new ArrayList<>();
        for (String key : keys) {
            Bson bson = Filters.eq(key, true);
            bsons.add(bson);
        }
        bsons.add(Filters.eq("_id", false));
        Bson filters = Filters.and(bsons.toArray(new Bson[]{}));
        FindIterable<Document> cursor = visit.find().projection(filters);
        List<String> dbs = new ArrayList<>();

        if (cursor != null) {
            for (Document document : cursor) {
                String[] notnulls = CglibUtils.NotNull.split(",");
                boolean isMatch = false;
                for (String notnull : notnulls) {
                    Object phone = document.get(notnull);
                    if (phone == null || phone.getClass() == BsonUndefined.class) {
                        isMatch = false;
                        break;
                    }
                    isMatch = true;
                }
                if (isMatch) {
                    String json = JSON.toJSONString(document, true);
                    dbs.add(json);
                }
            }
        }
        return dbs;
    }

}
