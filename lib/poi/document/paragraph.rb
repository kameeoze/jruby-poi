# -*- coding: utf-8 -*-
module POI
  class Paragraphs
    def initialize(document)
      @document = document
      @paragraphs = {}
    end
    
    def size
      @document._paragraphs.size
    end
    
    def [](index)
      @paragraphs[index] ||= Paragraph.new(@document.paragraph_array(index), @document)
    end
  end
  
  class Paragraph < Facade(:poi_paragraph, org.apache.poi.xwpf.usermodel.XWPFParagraph)
    def initialize(paragraph, document)
      @paragraph = paragraph
      @document = document
    end
    
    def poi_paragraph
      @paragraph
    end
    
    def text
      poi_paragraph ? poi_paragraph.getText() : "" 
    end
  end
end