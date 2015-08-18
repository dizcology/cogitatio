
class Nokogiri::XML::Node

# Finds and removes the pieces of text that are struck out from the XML tree
  def strike!
  
    # Finds all 'w:strike' tags in the XML tree that are descendants of this node
	self.xpath(".//w:strike").each do |s|

      # removes the text node containing the 'w:strike' tags
	  s.parent.parent.remove

    end

    # returns the cleaned up XML tree
	return self
  
  end
  
  # replaces 'original' with 'new' in the html representation of the XML tree
  def rep(original, new)
    # making sure we don't try substituting nil objects or for nil objects
	original=(original==nil)? "":original
    new=(new==nil)? "":new
    
	temp=self.inner_html
    self.inner_html=temp.gsub(original, new)
    self
  end

  def rows
    return nil unless self.is_tbl?
    
    rows=Array.new
    self.children.each do |r|
      next if r.name!="tr"
      rows << r
    end
    
    return rows
  end
  
  def cells
    return nil unless self.is_tr?
  
    cells=Array.new
    self.children.each do |c|
      next if c.name!="tc"
      cells << c
    end
    
    return cells
  end

  def is_tbl?
    self.name=="tbl"
  end
  
  def is_tr?
    self.name=="tr"
  end
  
  def add_onscr
    scrstyle=nil
    $pcs[7].each do |n|
      next if !(n.content.include?($type[7]))
      scrstyle=n.at(".//w:pStyle")

    end
    self.xpath(".//w:pStyle").each do |p|
      p["w:val"]="Onscreen"
    end
    
    self.xpath(".//w:rPr").each do |pr|
      pr << scrstyle.dup
    end
    self.xpath(".//w:pPr").each do |pr|
      pr << scrstyle.dup
    end
    self.xpath(".//w:tblPr").each do |pr|
      pr << $borders.dup
      pr << scrstyle.dup
    end
    
    self.xpath(".//w:r").each do |r|
      r.before("<w:rPr><w:pStyle w:val=\"Onscreen\"/></w:rPr>")
    end
    self.before("<w:pPr><w:pStyle w:val=\"Onscreen\"/></w:pPr>")
    
    self
  end
  
  def add_borders
    self.xpath(".//w:tblPr").each do |pr|
      pr << $borders
    end
    self
  end
  
  def remove_comments
    
    self.at(".//w:commentRangeStart").remove
    self.at(".//w:commentRangeEnd").remove
    self.at(".//w:commentReference").remove
  end
  
  def comment_start(id)
  
    self.before("<w:commentRangeStart w:id=\"#{id}\"/>")    
    self
  end

  def comment_end(id)
    self.after("<w:r><w:commentReference w:id=\"#{id}\"/></w:r>")
    self.after("<w:commentRangeEnd w:id=\"#{id}\"/>")
    
    self
  end
  
  def add_comment(id, cmt, author="Teacher", init="T")
  
    qq=$cmt.dup
    qq["w:id"]=id
    qq["w:author"]=author
    qq["w:date"]=DateTime.now.to_s
    qq["w:initials"]=init

    qq.rep("$cm",cmt)
    
    self.at(".//w:comments")<<qq
  end
  
  def merge_text_nodes
    prev_is_text = false

    newnodes = []
    self.children.each do |element|
      if element.text?
        if prev_is_text
          newnodes[-1].content += element.text
        else
          newnodes << element
        end
        element.remove
        prev_is_text = true
      else
        newnodes << element.merge_text_nodes
        element.remove
        prev_is_text = false
      end
    end

    self.children.remove  #this seems unnecessary
    newnodes.each do |item|
      self.add_child(item)
    end

    return self
  end
  
end
