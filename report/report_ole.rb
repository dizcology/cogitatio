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
    if self.match(/[ABC]/)
      abc=self.match(/([ABC])/)[1]
    end
    
    if self.include?("_")
      a, f = self.split("_")
      r="#{self} <- #{f}(#{a},na.rm=T)"
    elsif self.match(/^p[ABC]$/)

      r="#{self} <- fton(m[,\"%#{abc}\"])"
      
    elsif self.match(/^p\d{2}[ABC]$/)
      n=self.match(/^p(\d{2})[ABC]$/)[1]
      r="#{self} <- percent(p#{abc},#{n})"
    
    else
      r=nil
    end
      
    return r
  end
end

class Application < WIN32OLE
  def initialize(type)
    @type=type
    if @type=="docx"
      call_str='Word.Application'
      @state=0
    elsif @type=="pptx"
      call_str='PowerPoint.Application'
      @state=1
    else
      print "Unknown template type: #{@type}"
      exit
    end
    super(call_str)
    
    self.Visible=true
    self.activate
    self.WindowState=@state
    self.size(width=400,height=300)
    self.position(left=100,top=100)
    
  end
  
  def open(path)
    if @type=="docx"
      return self.Documents.Open(path)
    elsif @type=="pptx"
      return self.Presentations.Open(path)
    end
  end
  
  attr_reader :type
end

class WIN32OLE

  def size(width=400,height=300)
    self.Width=width
    self.Height=height
  end

  def position(left=0,top=0)
    self.Left=left
    self.Top=top
  end


  def each symb, &block
    count = self.send(symb).Count
    (1..count).each do |i|
      yield self.send(symb).Item({'index'=>i}) if block_given?
    end
  end
  
  def gsub(old,new)  
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
  
  def xgsub(old,new)  
    rng=self.UsedRange.Find(old)
    count=0
    if !(rng.nil?)
      begin
        rng.value=[[new]]
        count+=1
        rng=self.UsedRange.FindNext(rng)
      end until rng.nil?
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
  
  def pop_text(type)  #method for document or presentation objects
    if type=="pptx"
      self.each :Slides do |s|
        s.each :Shapes do |sh|
          next unless sh.HasTextFrame==-1
          sh.pop_tf if sh.TextFrame.HasText==-1
        end
      end
      
    elsif type=="docx"         
      $measures.each do |m|
        next if m.type=="system" || m.value.nil? || m.type=="Type"
        self.Application.gsub(m.tag,m.value.round(2).to_s)
      end
    end
  end
  
  def pop_shape(type) #method for document or presentation objects
    if type=="pptx"
      self.each :slides do |s|
        s.each :Shapes do |sh|
            sh.pop_chart if sh.HasChart==-1
            sh.pop_table if sh.HasTable==-1
            #sh.pop_tf if sh.TextFrame.HasText==-1

        end
      end
    elsif type=="docx"
      self.each :InlineShapes do |sh|
        sh.pop_chart
      end
    end
  end
  
  def pop_tf  #method for shape objects
    $measures.each do |m|
      next if m.type=="system" || m.value.nil? || m.type=="Type"
      txt=self.TextFrame.TextRange.Text.gsub(m.tag,m.value.round(2).to_s)
      self.TextFrame.TextRange.Text=txt
    end
  end
  
  def pop_table #method for shape objects
    tbl=self.Table
    tbl.each :Rows do |r|
      r.each :Cells do |c|
        next if c.Shape.TextFrame.HasText!=-1
        $measures.each do |m|
          next if m.type=="system" || m.value.nil? || m.type=="Type"
          txt=c.Shape.TextFrame.TextRange.Text.gsub(m.tag,m.value.round(2).to_s)
          c.Shape.TextFrame.TextRange.Text=txt
        end
      end
    end
  end
  
  def pop_chart #method for shape objects
    cd=self.Chart.ChartData
    cd.activate
    wrksht=cd.Workbook.Worksheets(1)
    
    $measures.each do |m|
      next if m.type=="system" || m.value.nil? || m.type=="Type"
      wrksht.xgsub(m.tag,m.value.round(2).to_s)
      
      txt=self.Chart.ChartTitle.Text.gsub(m.tag,m.value.round(2).to_s)
      self.Chart.ChartTitle.Text=txt
      
      txt=self.Chart.Axes({'Type'=>2}).AxisTitle.Text.gsub(m.tag,m.value.round(2).to_s)
      self.Chart.Axes({'Type'=>2}).AxisTitle.Text=txt
      
      #HERE
      #self.Chart.each :Axes do |a|
      #  sh.Chart.Axes({'Type'=>2}).AxisTitle.Text
      #  txt1=a.Axistitle.Caption.gsub(m.tag,m.value.round(2).to_s)
      #  a.Axistitle.Caption=txt1
      #end
    end

    cd.Workbook.Close({'SaveChanges'=>'True'})
    
    sleep 0.3 #temporary solution, want asynchronosity
    
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
    @value=nil
    #@tag="$#{@mid}"  #get this from measures_template.csv
    self.get_req_string
  end
  
  def get_req_string
    @req=@mid.req
  end
  
  def get_value
  
    if !(@value.nil?)
      #do nothing if there is already a value
    elsif @mid.match("_")
      anc=@mid.split("_")[0]
      ancestor=$measures.select{|a| a.mid==anc}[0]
      if ancestor.nil? || ancestor.req.nil?
        @value=nil
      else
        R.eval(@req)  #returns true if successful
        @value=R.pull("as.numeric(#{@mid})")  #Kernel.eval("R.#{@mid}") 
      end
    elsif @req.nil?
      @value=nil
    else
      R.eval(@req)  #returns true if successful
      @value=R.pull("as.numeric(#{@mid})")  #Kernel.eval("R.#{@mid}") 
    end
    
  end
  
  attr_accessor :mid, :tag, :value, :description, :req, :type, :alias

end

$measures = Array.new

def $measures.dump(pth)
  begin
    f=File.open(pth,"w")
  rescue
    print "Can't create file: "+pth
    exit
  end
  
  self.each do |m|
    val=(m.value.class==Array)? "*":m.value  #arrays are not printed
    f.print [m.mid,m.tag,val,m.type,m.alias,m.description].join(",")+"\n"
  
  end
end

testing=true

if testing==true

  $RUNPATH="C:\\Users\\yliu\\SkyDrive\\RM-synced\\cogitatio\\report\\"
  $PATH=$RUNPATH #"C:\\Users\\yliu\\SkyDrive\\RM-synced\\ANALYSIS REPORT\\"
  Dir.chdir($PATH)
  metric_path=$RUNPATH+"metrics\\Metrics_Report.csv"
  template_path=$RUNPATH+"templates\\2013-2014 Data Reporting Template tagged.pptx"
  out_path = $RUNPATH+"outputs\\"
else

  tle1="Open metric report file."
  puts tle1
  metric_path=getfilepath(tle1)

  #this depends on the shape of the APEX output
  #TODO: use school/district name instead
  
  exit if metric_path==""

  tle2="Open template file."
  puts tle2
  template_path=getfilepath(tle2)
  exit if template_path==""

end

metric_path_R="\""+metric_path.gsub("\\","/")+"\""
metric_name=metric_path.split("\\")[-1].split(".")[0]
template_type=template_path.split(".")[-1]

puts template_type
puts metric_name

exit

source_R="\"C:/Users/yliu/SkyDrive/RM-synced/cogitatio/report/agg.r\""
preamble = <<EOF
  source(#{source_R})
  m0 <- read.csv(#{metric_path_R},head = TRUE, sep = ",")
  m <- m0[3:dim(m0)[1],]
  colnames(m)=as.vector(as.matrix(m0[1,]))
EOF
R.eval(preamble)


metric=File.open(metric_path,"r")

mf=File.open(template_path+"measures_template.csv")

mf.each do |line|
  #the Measure object with mid="MID" records the header row of measures.csv
  mid, tag, value, type, als, description = line.strip.split(",")
  newm = Measure.new(mid)
  newm.tag=tag
  newm.value= (value.strip=="")? nil : value
  newm.type=type
  newm.alias=als
  newm.description=description
  
  newm.get_value
  
  $measures << newm
end

app=Application.new(template_type)
temp=app.open(template_path)

=begin
temp.each :Slides do |s|
  s.each :Shapes do |sh|
    if sh.HasChart==-1
      #puts sh.Chart.Name
      
      #puts sh.Chart.Axes.Count
      
      puts sh.Chart.Axes({'Type'=>2}).AxisTitle.Text #if sh.Chart.Axes({'Type'=>2}).HasTitle==-1
      
      #sh.Chart.each :Axes do |a|
      #  print a#.AxisTitle.Caption if a.HasTitle==-1
      #end
      #gets
    end
    
  end
end

exit
=end


temp.pop_text(template_type)
temp.pop_shape(template_type)

exit

temp.SaveAs($PATH+metric_name+"."+template_type)
$measures.dump($PATH+metric_name+"_measures.csv")

temp.Application.size(1000,1000)

