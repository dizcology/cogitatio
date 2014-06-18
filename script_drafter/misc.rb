
class String
  def content
    self
  end
  
  def rep(original, new)
    return self unless self.include?(original)   
    
    self.gsub!(original, new)
    self
  end
end

class NilClass
  def content
    return ""
  end
  
  def upcase
    return 1
  end
  
  def [](n)
    return ""
  end
  
  def remove
    self
  end
end

class Array
  def add_new(add)
    return self if add==self[-1]
    
    self << add
    return self
  end
end
