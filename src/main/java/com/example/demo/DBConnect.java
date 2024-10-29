package com.example.demo;

import java.sql.*;

public class DBConnect {
    private static final String DRIVER = "oracle.jdbc.OracleDriver";
    private static final String URL = "jdbc:oracle:thin:@192.168.1.93:1521:data03";   // jdbc:oracle:thin: @主机地址 :  端口号 : 实例名
    private static final String USER = "fcl";
    private static final String PASSWORD = "fcl";
    static {//因为驱动只需要加载一次，所以在静态语句中进行
        try {
            Class.forName(DRIVER);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    static String aSQL1 = "select"
            + " FUND.ST_FUNDNAMELONG,"
            + " FUND.ST_FUNDNAME,"
            + " ITCOMPANY.ST_ITCOMPANY,"
            + " ASSMALLCLASS.ST_ASSMALLCLASS,"
            + " EXCHANGE.ST_EXCHANGE,"
            + " FUND.RA_TARGET,"
            + " BUNPAI.ST_BUNPAI,"
            + " INDEXNAME1.ST_INDEXNAME,"
            + " FUND.ST_BUY,"
            + " FUND.ST_REALIZE,"
            + " FUND.RA_STOCKUPPER,"
            + " FUND.RA_FOREIGNUPPER,"
            + " FUND.DT_ESTABLISH,"
            + " FUND.SM_ESTABLISH,"
            + " FUND.DT_REPAY,"
            + " DECODE(FUND.BO_NOLIMIT,'0','いいえ','1','はい'),"
            + " FUND.SM_SELLFEE,"
            + " SLIDEFEE.ST_SLIDEFEE,"
            + " FUND.SM_CANCELFEE,"
            + " DECODE(FUND.FG_RESERVATION, '0','なし','1','％（解約時）','2','￥（解約時）','3','％（追加設定時）','4','￥（追加設定時）','5','％（追加/解約時）','6','￥（追加/解約時）','7','％（途中解約時）','8','￥（途中解約時）') ,"
            + " FUND.SM_RESERVATION,"
            + " FAVORABLE.ST_FAVORABLE,"
            + " BANK.ST_BANKKANJI"
            + " from"
            + " FUND FUND,"
            + " ITCOMPANY ITCOMPANY,"
            + " SMALLCLASS SMALLCLASS,"
            + " ASSMALLCLASS ASSMALLCLASS,"
            + " EXCHANGE EXCHANGE,"
            + " BUNPAI BUNPAI,"
            + " INDEXNAME INDEXNAME1,"
            + " FAVORABLE FAVORABLE,"
            + " SLIDEFEE SLIDEFEE,"
            + " BANK BANK,"
            + " FUND2 FUND2"
            + " where"
            + " FUND.ID_FUND = FUND2.ID_FUND and"
            + " FUND.ID_ITCOMPANY = ITCOMPANY.ID_ITCOMPANY and"
            + " FUND.ID_SMALLCLASS = SMALLCLASS.ID_SMALLCLASS and"
            + " FUND.ID_ASSMALLCLASS = ASSMALLCLASS.ID_ASSMALLCLASS and"
            + " FUND.ID_EXCHANGE = EXCHANGE.ID_EXCHANGE and"
            + " FUND.ID_BUNPAI = BUNPAI.ID_BUNPAI and"
            + " FUND.ID_INDEXNAME = INDEXNAME1.ID_INDEXNAME and"
            + " FUND.ID_FAVORABLE = FAVORABLE.ID_FAVORABLE and"
            + " FUND.ID_SLIDEFEE = SLIDEFEE.ID_SLIDEFEE and"
            + " FUND.ID_TRUSTBANK = BANK.ID_BANK and"
            + " FUND.ID_FUND = '79312024'";

    String sql2 = "select"
            + " FUND.ID_FUND,"
            + " FUND.ST_FUNDNAMELONG,"
            + " SETTLINGDAY.ST_SETTLINGDAY"
            + " from"
            + " FUND,"
            + " SETTLINGDAY"
            + " where"
            + " FUND.ID_FUND = SETTLINGDAY.ID_FUND"
            + " and"
            + " FUND.ID_FUND = '79312024'"
            + " ORDER BY ST_SETTLINGDAY";

    //信託報酬
    String sql3 = "select"
            + " FUND.ID_FUND,"
            + " FUND.ST_FUNDNAMELONG ,"
            + " REWARD.ST_CONDITION ,"
            + " DECODE(REWARD.ID_REWARD,'0','合計','1','委託会社','2','販売会社','3','受託会社'),"
            + " REWARD.RA_REWARD"
            + " from"
            + " FCL.FUND FUND,"
            + " FCL.REWARD REWARD"
            + " where"
            + " FUND.ID_FUND = REWARD.ID_FUND"
            + " and"
            + " FUND.ID_FUND = '79312024'"
            + " ORDER BY ID_REWARD ";
    //運用方針
    String sql4 = "select"
            + " FUND.ID_FUND,"
            + " FUND.ST_FUNDNAMELONG,"
            + " COURSE.NO_TOPIC,"
            + " COURSE.ST_COURSE"
            + " from"
            + " FCL.FUND FUND,"
            + " FCL.COURSE COURSE"
            + " where"
            + " FUND.ID_FUND = COURSE.ID_FUND"
            + " and"
            + " FUND.ID_FUND = '79312024' and NO_ROWS = '1'";
    //販売会社
    static String sql5 = "select"
            + " FUND.ID_FUND ,"
            + " FUND.ST_FUNDNAMELONG ,"
            + " BANK.ST_BANKKANJI"
            + " from"
            + " FCL.FUND FUND,"
            + " FCL.SELLBANK SELLBANK,"
            + " FCL.BANK BANK"
            + " where"
            + " FUND.ID_FUND = SELLBANK.ID_FUND and"
            + " SELLBANK.ID_BANK = BANK.ID_BANK"
            + " and"
            + " FUND.ID_FUND = '79312024'"
            + " ORDER BY BANK.ST_BANKKANA";
    static String sql6 = "select FUND.ID_FUND, FUND.ST_FUNDNAME,FUND.ST_FUNDNAMELONG     from FUND FUND     where     FUND.ST_FUNDNAME like ";


    /**
     * 建立与数据库的连接
     *
     * @return 连接好的连接
     */
    public static Connection getConnection() {
        try {
            Connection connection = DriverManager.getConnection(URL, USER, PASSWORD);
            return connection;
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return null;
    }

    public static String[] queryArray(String sql) {
        Connection connection = null;
        PreparedStatement preparedStatement = null;
        ResultSet resultSet = null;
        String[] resultArray = new String[0];
        try {
            //获取连接
            connection = getConnection();
            //创建语句对象
            preparedStatement = connection.prepareStatement(sql);
            //执行SQL语句
            resultSet = preparedStatement.executeQuery();
            ResultSetMetaData metaData = resultSet.getMetaData();
            resultArray = new String[0];
            while (resultSet.next()) {            //行数据
                String[] newArray = new String[resultArray.length + 1];
                System.arraycopy(resultArray, 0, newArray, 0, resultArray.length);
                newArray[resultArray.length] = resultSet.getString("columnName");
                resultArray = newArray;
            }

        } catch (SQLException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            close(resultSet, preparedStatement, connection);
        }
        return resultArray;
    }

    public static void close(ResultSet resultSet, Statement statement, Connection connection) {//查询时
        try {
            if (resultSet != null && !resultSet.isClosed())
                resultSet.close();
        } catch (SQLException ex) {
            ex.printStackTrace();
        } finally {
            try {
                if (statement != null && !statement.isClosed())
                    statement.close();
            } catch (SQLException ex) {
                ex.printStackTrace();
            } finally {
                try {
                    if (connection != null && !connection.isClosed())
                        connection.close();
                } catch (SQLException ex) {
                    ex.printStackTrace();
                }
            }
        }
    }
}