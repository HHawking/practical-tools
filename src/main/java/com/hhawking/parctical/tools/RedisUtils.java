package com.hhawking.parctical.tools;

import redis.clients.jedis.Jedis;
import redis.clients.jedis.JedisPool;
import redis.clients.jedis.JedisPoolConfig;

import java.util.ResourceBundle;

public class RedisUtils {

    private static JedisPool jedisPool;

    private RedisUtils() {}

    static {
        ResourceBundle rb = ResourceBundle.getBundle("jedis");
        int maxtotal = Integer.parseInt(rb.getString("maxtotal"));
        int maxwaitmillis = Integer.parseInt(rb.getString("maxwaitmillis"));
        int timeout = Integer.parseInt(rb.getString("timeout"));
        int maxIdle = Integer.parseInt(rb.getString("maxIdle"));
        String host = rb.getString("host");
        int port = Integer.parseInt(rb.getString("port"));
        String password = rb.getString("password");

        JedisPoolConfig config = new JedisPoolConfig();

        //控制一个pool可分配多少个jedis实例，通过pool.getResource()来获取；
        //如果赋值为-1，则表示不限制；如果pool已经分配了maxActive个jedis实例，则此时pool的状态为exhausted(耗尽)。
        config.setMaxTotal(maxtotal);

        //控制一个pool最多有多少个状态为idle(空闲的)的jedis实例。
        config.setMaxIdle(maxIdle);

        //表示当borrow(引入)一个jedis实例时，最大的等待时间，如果超过等待时间，则直接抛出JedisConnectionException；
        config.setMaxWaitMillis(maxwaitmillis);

        //在borrow一个jedis实例时，是否提前进行validate操作；如果为true，则得到的jedis实例均是可用的；
        config.setTestOnBorrow(true);

        jedisPool = new JedisPool(config, host,port, timeout,password);
    }

    /**
    * 从jedis连接池中获取获取jedis对象
    * @return
    */
    public static Jedis getJedis() {
        return jedisPool.getResource();
    }

}
