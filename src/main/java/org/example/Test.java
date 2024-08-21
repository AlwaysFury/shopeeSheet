package org.example;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.math.BigInteger;

/**
 * @Description
 * @Author 罗宇航
 * @Date 2024-4-5
 */
public class Test {
    public static void main(String[] args) {
        shop_auth();
    }

    public static void shop_auth(){
        long timest = System.currentTimeMillis() / 1000L;
        String host = "https://partner.shopeemobile.com";
        String path = "/api/v2/shop/auth_partner";
        String redirect_url = "https://www.baidu.com/";
        long partner_id = 1140406;
        String tmp_partner_key = "6a57414d766c79786863726c58624f466b51525673485375644756436f4b6c74";
        String tmp_base_string = String.format("%s%s%s", partner_id, path, timest);
        byte[] partner_key;
        byte[] base_string;
        String sign = "";
        try {
            base_string = tmp_base_string.getBytes("UTF-8");
            partner_key = tmp_partner_key.getBytes("UTF-8");
            Mac mac = Mac.getInstance("HmacSHA256");
            SecretKeySpec secret_key = new SecretKeySpec(partner_key, "HmacSHA256");
            mac.init(secret_key);
            sign = String.format("%064x",new BigInteger(1,mac.doFinal(base_string)));
        } catch (Exception e) {
            e.printStackTrace();
        }
        String url = host + path + String.format("?partner_id=%s&timestamp=%s&sign=%s&redirect=%s", partner_id,timest, sign, redirect_url);
        System.out.println(url);
    }
}
