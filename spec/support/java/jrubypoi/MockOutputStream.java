package jrubypoi;
import java.io.*;

public class MockOutputStream extends OutputStream {
  public String name;
  public boolean write_called;
  
  public MockOutputStream(String name) {
    this.name = name;
    this.write_called = false;
  }
  
  public void write(int b) throws IOException {
    this.write_called = true;
  }

  public void write(byte[] b) throws IOException {
    this.write_called = true;
  }

  public void write(byte[] b, int offset, int length) throws IOException {
    this.write_called = true;
  }
}