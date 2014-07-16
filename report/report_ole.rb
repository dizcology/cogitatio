require 'rubygems'
require 'zip/zip' # rubyzip gem
require 'win32ole'
require 'fileutils'
require 'io/console'
require 'Date'
require_relative 'dialog_boxes_report.rb' 
require 'rinruby'

class String

  def mesh(ary, sep="")
    a=Array.new
    ary.each do |x|
      a << self+sep+x.to_s
    end
    return a
  end
  
  def req #method to generate the correct R code strings
    r=""
    abc=self.match(/([ABC])/)[1]
    
    
    if self.include?("_")
      a, f = self.split("_")
      r="#{self} <- #{f}(#{a},na.rm=T)"
    elsif self.match(/^p[ABC]$/)

      r="#{self} <- fton(m[,\"%#{abc}\"])"
      
    elsif self.match(/^p\d{2}[ABC]$/)
      n=self.match(/^p(\d{2})[ABC]$/)[1]
      r="#{self} <- percent(p#{abc},#{n})"
    end
      
    return r
  end
end



class WIN32OLE

  def size(width=400,height=300)
    self.Width=width
    self.Height=height
  end

  def position(left=0,top=0)
    self.Move({'Left'=>left,'Top'=>top})
  end


  def gsub(old,new)  #define in WIN32OLE class
    self.Selection.HomeKey(unit=6)
    find=self.Selection.Find
    find.Text=old
    count=0
    while find.Execute
      self.Selection.TypeText(text=new)
      count+=1
    end
    return count
  end

  def insert(tag,img=kitten, scale=100, replace=false)
    self.Selection.HomeKey(unit=6)
    find=self.Selection.Find
    find.Text=tag
    find.Execute
    
    if replace
      
      self.Selection.TypeText(text="\n")
      self.Selection.Move({'Unit'=>1,'Count'=>-1})
    else
      self.Selection.Collapse
      self.Selection.TypeText(text="\n")
      self.Selection.Move({'Unit'=>1,'Count'=>-1})
    end
    range=self.Selection.Range
    #range.Start-=1
    #range.End-=1
    pic=range.InlineShapes.AddPicture(img)
    pic.ScaleHeight=scale
    pic.ScaleWidth=scale
  end

  
  def insertchart(tag, type, replace=false)
    #list of char types: 
    #http://msdn.microsoft.com/en-us/library/ff838409(v=office.14).aspx
    self.Selection.HomeKey(unit=6)
    find=self.Selection.Find
    find.Text=tag
    find.Execute
    
    if replace
      self.Selection.TypeText(text="\n")
      self.Selection.Move({'Unit'=>1,'Count'=>-1})
    else
      self.Selection.Collapse
      self.Selection.TypeText(text="\n")
      self.Selection.Move({'Unit'=>1,'Count'=>-1})
    end
    
    cht=self.Selection.InlineShapes.AddChart({'Type'=>type})

  end
end


class Object
  def in?(ary)
    return ary.include?(self)
  end
end

class Measure
  def initialize(str)
    @mid=str
    @tag="$#{@mid}"
    self.req_string
  end
  
  def req_string
    @req=@mid.req
  end
  
  def get_value
    R.eval(@req)  #returns true if successful
    @value=R.pull(@mid)  #Kernel.eval("R.#{@mid}")  
  end
  
  attr_accessor :mid, :tag, :value, :description, :req, :type

end

$PATH="C:\\Users\\yliu\\SkyDrive\\RM-synced\\ANALYSIS REPORT\\"
Dir.chdir($PATH)


tle1="Open metric report file."
puts tle1
metric_path=getfilepath(tle1)
metric_path_R="\""+metric_path.gsub("\\","/")+"\""

source_R="\"C:/Users/yliu/SkyDrive/RM-synced/cogitatio/report/agg.r\""
preamble = <<EOF
  source(#{source_R})
  m0 <- read.csv(#{metric_path_R},head = TRUE, sep = ",")
  m <- m0[3:dim(m0)[1],]
  colnames(m)=as.vector(as.matrix(m0[1,]))
EOF
R.eval(preamble)


metric=File.open(metric_path,"r")
#metric.readline #row with years

desc_stats=["mean","sd","min","max"]
system=["pA","pB","pC"]
aggregated=["p75A", "p45B", "p30C"]

system.each do |m|
  aggregated|=m.mesh(desc_stats,"_")
end

list=Array.new

(system|aggregated).each do |m|
  list << Measure.new(m)
  list[-1].get_value
end

word=WIN32OLE.new('Word.Application')
word.Visible=true
doc=word.Documents.Open($PATH+"template.docx")
word.activate
word.WindowState=0

word.size(width=400,height=300)
word.position(left=100,top=100)

kitten="C:\\Users\\yliu\\Desktop\\kitten.jpg"


list.each do |m|
  next if m.mid.in?(system)
  
  word.gsub(m.tag,m.value.round(1).to_s)

end

tag1="Table"
tag2="CHART"
pic=word.insert(tag1,img=kitten,scale=50, replace=false)

chrt=word.insertchart(tag2, type=51, replace=true).Chart
chrt.ChartData.Activate
wrksht=chrt.ChartData.Workbook.Worksheets(1)
#puts wrksht.ole_methods
#gets
#puts chrt.SeriesCollection.ole_methods
#puts chrt.SeriesCollection.count
#puts chrt.SeriesCollection(1).ole_methods
#gets
#exit

#puts wrkbk.Worksheets(1).UsedRange.ClearContents
#print wrkbk.Worksheets(1).Range("B1:B2").value
wrksht.Range("A2:A4").value=[["A"],["B"],["C"]]
wrksht.Range("B2:B7").value=[[list[4].value],[list[3].value],[list[4].value],[list[3].value],[list[4].value],[list[3].value]]
puts chrt.SeriesCollection.count
chrt.SeriesCollection.NewSeries
wrksht.Range("E1:E2").value=[["NAME"],[list[3].value]]
puts chrt.SeriesCollection(1).name="WOOPIE!"


exit
doc.SaveAs($PATH+"out_ole.docx")
#word.Activate
#ord.WindowState=1
exit
