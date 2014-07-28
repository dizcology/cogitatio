require 'rubygems'
require 'zip/zip' # rubyzip gem
require 'nokogiri'
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

class Nokogiri::XML::Node
  
  def rep(original, new)
    original=(original==nil)? "":original
    new=(new==nil)? "":new
    temp=self.inner_html
    self.inner_html=temp.gsub(original, new)
    self
  end

end

class DOCX
  def self.open(path)
    self.new(path)
  end


  def initialize(path)
    @replace = {}
    @zip = Zip::ZipFile.open(path)
    @doc=Nokogiri::XML(@zip.read("word/document.xml")) {|x| x.noent}
  end

  attr_accessor :doc, :replace

  def save(path)
    
    @replace["word/document.xml"] = @doc.to_xml
    
    Zip::ZipFile.open(path, Zip::ZipFile::CREATE) do |output|
      @zip.each do |entry|
        output.get_output_stream(entry.name) do |o|
          if @replace[entry.name]
            o.write(@replace[entry.name])
          else
            o.write(@zip.read(entry.name))
          end
        end
      end
    end
    @zip.close
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

ff=DOCX.open("template.docx")

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

list.each do |measure|  #"$pA" appears also in "$pA_mean"
  next if system.include?(measure.mid)
  
  ff.doc.rep(measure.tag,measure.value.round(1).to_s)
  
end

ff.save("out.docx")

exit
