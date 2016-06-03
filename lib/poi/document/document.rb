# -*- coding: utf-8 -*-
module POI
  class Document < Facade(:poi_document, org.apache.poi.xwpf.usermodel.XWPFDocument)
    def self.open(filename_or_stream)
       name, stream = if filename_or_stream.kind_of?(java.io.InputStream)
         [File.join(Dir.tmpdir, "worddoc.docx"), filename_or_stream]
       elsif filename_or_stream.kind_of?(IO) || StringIO === filename_or_stream || filename_or_stream.respond_to?(:read)
         # NOTE: The String.unpack hear can be very inefficient on Large files
         [File.join(Dir.tmpdir, "worddoc.docx"), java.io.ByteArrayInputStream.new(filename_or_steram.read.unpack('c*').to_java(:byte))]
       else
         raise Exception, "FileNotFound" unless File.exists?(filename_or_stream)
         [filename_or_stream, java.io.FileInputStream.new(filename_or_stream)]
       end
       instance = self.new(name, stream)
       if block_given?
         result = yield instance
         return result
       end
       instance
    end
    
    attr_reader :filename
    
    def initialize(filename, io_stream, options={})
      @filename = filename
      @docfile = if io_stream
        org.apache.poi.xwpf.usermodel.XWPFDocument.new(io_stream)
      elsif options[:format] == :poifs
        org.apache.poi.poifs.filesystem.POIFSFileSystem.new(io_stream)
      else
        org.apache.poi.xwpf.usermodel.XWPFDocument.new
      end
    end
    
    def poi_document
      @docfile
    end
    
    alias :_paragraphs :paragraphs
    def paragraphs
      @paragraphs ||= Paragraphs.new(self)
    end
    
    alias :_tables :tables
    def tables
      @tables ||= Tables.new(self)
    end
  end
end