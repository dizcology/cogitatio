class ExpressionsController < ApplicationController
  def index
    #@result = []
    #@input = ""
  end

  def create
    #fail
    input = params["expression"]["input"]
    session[:input]
    session[:result] = generate(input, params["expression"]["leading_plus"])
    redirect_to "/"
  end
  
  #private
  
  def generate(str, lp)
    hsh = {}
    divide = /(?:[+-]|\A)\w+/
    
    terms = str.scan(divide)
    terms.each do |term|
      hsh[term] = reps(term.gsub(/\s/,""))  
    end
     
    result = assemble(hsh)
    
    #generate cases where /\A\+/ is removed
    result = redundant(result)
    
    if lp == "0"
      #remove leading +
      result = result.map{|x| x.sub(/\A\+/,"")}
    end
    
    return result.uniq
  end
  
  def redundant(a)
    b = []
    a.each do |x|
      b << x
      if x.match(/\A\+/)
        b << x.sub("+","")
      end
    end
    return b.sort
  end
  
  
  def assemble(h)
    if h.keys.size == 1
      return h[h.keys[0]].sort
    end
    
    result = []
    h.keys.each do |this_key|
      this_array = h[this_key]
      comp_hash = h.select{|k,v| k!= this_key}
      comp_array = assemble(comp_hash)
      result += this_array.product(comp_array).collect{|x,y| x+y}
    end
    
    return result.sort
  end
  
  def reps(term) 
    mono = term.gsub(/[+-]/,"").gsub("\W","")
    temp = swaps(mono)
    if term[0] == "-"
      temp = temp.map{|t| "-"+t}
    else
      temp = temp.map{|t| "+"+t}
    end
    
    return temp
  end
  
  def swaps(mon)
    if !mon.match(/[^0-9]/) || !mon.match(/[^a-zA-Z]/)
      return [mon]
    end
    
    a, b = mon.match(/(\d+)([^0-9])/)[1..2]
    return [a+b,a+"*"+b,b+"*"+a]
  end
end
