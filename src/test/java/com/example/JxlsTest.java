package com.example;

import static org.testng.Assert.fail;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;

import org.jxls.common.Context;
import org.jxls.jdbc.JdbcHelper;
import org.jxls.util.JxlsHelper;
import org.testng.annotations.Test;

public class JxlsTest {

  @Test
  public void export() {
    try {
      Connection conn = DriverManager.getConnection("jdbc:h2:mem:test", "sa", "");
      Context context = new Context();
      JdbcHelper helper = new JdbcHelper(conn);

      context.putVar("jdbc", helper);

      InputStream input = getClass().getResourceAsStream("test.xls");
      OutputStream output = new ByteArrayOutputStream();

      JxlsHelper.getInstance().setEvaluateFormulas(true).processTemplate(input, output, context);

      conn.close();
    } catch (Exception e) {
      fail("", e);
    }
  }
}
