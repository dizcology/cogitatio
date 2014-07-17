def show(msg,tle,type)
  puts msg
  showmessage(msg,tle,type)
end

def getfolder(msg)
  puts msg
  getfolderpath(msg)
end

def getfile(msg)
  puts msg
  getfilepath(msg)
end
class String
  def content
    self
  end
  
  def rep(original, new)
    return self unless self.include?(original)   
    
    self.gsub!(original, new)
    self
  end
  
  def stage?
    $stage_names.each do |sn|
      if self.include?(sn)
        return true
      end
    end
    
    return false
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
  
  def strike!
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
