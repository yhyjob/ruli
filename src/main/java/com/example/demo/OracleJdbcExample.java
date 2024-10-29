package com.example.demo;
import java.sql.*;

public class OracleJdbcExample {
    public static void main (String args []) throws SQLException, ClassNotFoundException {
        // Load Oracle driver
        //DriverManager.registerDriver (new oracle.jdbc.driver.OracleDriver());
        Class.forName("oracle.jdbc.OracleDriver");

        // Connect to the local database
        Connection conn =
                DriverManager.getConnection ("jdbc:oracle:thin:@192.168.1.93:1521:data03",
                        "fcl", "fcl");

        // Query the employee names
        Statement stmt = conn.createStatement ();
        ResultSet rset = stmt.executeQuery ("select ename from emp");

        // Print the name out
        while (rset.next ())
            System.out.println (rset.getString (1));
    }
}

